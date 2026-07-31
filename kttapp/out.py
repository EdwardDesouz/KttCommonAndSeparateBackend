from django.db import connections,connection
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from django.http import HttpResponse, JsonResponse
from django.http import FileResponse
from django.conf import settings
from django.db import transaction
from datetime import datetime,timedelta
from openpyxl import load_workbook
from reportlab.platypus import Table, TableStyle
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.graphics.barcode import code39
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.utils import simpleSplit
from .utils import to_float, get_val
from PyPDF2 import PdfReader, PdfWriter
from xml.etree.ElementTree import Element, SubElement, tostring
from .views import SqlDb
import xlwt
import pandas as pd
import zipfile
import logging
import json
import traceback
import os
import io



class OutList(APIView):
    def get(self, request):
        try:
            Username = request.query_params.get("user") or request.session.get("Username")
            if not Username:
                return Response({"error": "Username required"}, status=400)

            # Get logged in user's MailBoxId
            account_rows = SqlDb.execute_query(
                "SELECT AccountId, MailBoxId FROM ManageUser WHERE UserName = %s",
                [Username]
            )
            if not account_rows:
                return Response({"error": "User not found"}, status=404)

            MailBoxId = account_rows[0]["MailBoxId"]

            show_all = request.query_params.get("all", "false").lower() == "true"

            base_select = """
                SELECT
                    t1.Id AS ID,
                    t1.JobId,
                    t1.PermitId,
                    t1.MSGId,
                    CONVERT(varchar, t1.TouchTime, 105) AS DECDATE,
                    SUBSTRING(t1.DeclarationType, 1, CHARINDEX(':', t1.DeclarationType) - 1) AS DECTYPE,
                    t1.TouchUser AS CREATE_USER,
                    t1.TradeNetMailboxID AS DECID,
                    CONVERT(varchar, t1.DepartureDate, 105) AS ETA,
                    t1.PermitNumber AS PERMITNO,
                    i.Name + ' ' + i.Name1 AS EXPORTER,

                    -- HAWB: aggregate all item-level OutHAWBOBL values for this permit
                    -- (mirrors legacy outListTable behaviour, which pulled this from
                    -- OutItemDtl; here we pull it from CommonItemDtl instead)
                    STUFF((
                        SELECT DISTINCT ', ' + ci.OutHAWBOBL
                        FROM CommonItemDtl ci
                        WHERE ci.PermitId = t1.PermitId
                          AND ci.OutHAWBOBL IS NOT NULL
                          AND ci.OutHAWBOBL != ''
                        FOR XML PATH('')
                    ), 1, 2, '') AS HAWB,

                    CASE
                        WHEN t1.OutwardTransportMode = '4 : Air' THEN t1.OutMasterAirwayBill
                        WHEN t1.OutwardTransportMode = '1 : Sea' THEN t1.OutOceanBillofLadingNo
                        ELSE ''
                    END AS MAWBOBL,

                    t1.LoadingPortCode AS POL,

                    -- Previously missing fields the frontend columns config expects
                    t1.DischargePort AS POD,
                    CASE
                        WHEN t1.COType = '--Select--' THEN ''
                        ELSE ISNULL(t1.COType, '')
                    END AS COTYPE,
                    CASE
                        WHEN t1.CerDetailtype1 = '--Select--' THEN ''
                        ELSE ISNULL(t1.CerDetailtype1, '')
                    END AS CERTTYPE,
                    t1.CertificateNumber AS CERTNO,

                    t1.MessageType AS MSGTYPE,
                    t1.OutwardTransportMode AS TPT,
                    t1.PreviousPermit AS PREPMT,
                    t1.GrossReference AS XREF,
                    t1.InternalRemarks AS INTREM,
                    t1.Message As MSG,
                    t1.TotalGSTTaxAmt AS GSTAMT,
                    t1.Status,

                    -- REL / RCL: inferred as ReleaseLocation / RecepitLocation
                    -- (confirm with the frontend team if this mapping is wrong)
                    t1.ReleaseLocation AS REL,
                    t1.RecepitLocation AS RCL,

                    CASE
                        WHEN t1.Status = 'APR' THEN
                            CASE
                                WHEN EXISTS (
                                    SELECT 1 FROM CommonPMT p
                                    WHERE p.PermitNumber = t1.PermitNumber
                                    AND p.ConditionCode IN ('Z02','Z18','Z06')
                                ) THEN 'RED'
                                WHEN EXISTS (
                                    SELECT 1 FROM CommonPMT p
                                    WHERE p.PermitNumber = t1.PermitNumber
                                    AND p.ConditionCode IN ('D6','D3')
                                ) THEN 'MAROON'
                                ELSE 'DEFAULT'
                            END
                        ELSE 'DEFAULT'
                    END AS COLOR
                FROM CommonHeaderTbl t1
                LEFT JOIN CommonExporter i ON t1.ExporterCompanyCode = i.Code
            """

            if show_all:
                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'OUTDEC'
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId])

            else:
                nowdate = datetime.now() - timedelta(days=90)
                date_filter = nowdate.strftime("%Y/%m/%d")

                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'OUTDEC'
                    AND CONVERT(varchar, t1.TouchTime, 111) >= %s
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId, date_filter])

            return Response(result)

        except Exception as e:
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)
# New Permit
class OutNewPermit(APIView):
    def get(self, request):
        try:
            Username = request.query_params.get("user")
            if not Username:
                Username = request.session.get("Username")
            if not Username:
                return Response({"error": "Session expired or User not provided"}, status=401)
            refDate = datetime.now().strftime("%Y%m%d")
            yy_mmdd = datetime.now().strftime("%Y-%m-%d")
            currentDate = datetime.now().strftime("%d/%m/%Y")  

            q_account = "SELECT AccountId FROM ManageUser WHERE UserName = %s"

            account_rows = SqlDb.execute_query(q_account, [Username])
            if not account_rows:
                return Response({"error": "User not found"}, status=404)
            AccountId = account_rows[0]['AccountId']


# 6-6-2026 Start
            # ref_rows = SqlDb.execute_query(
            #     """
            #     SELECT ISNULL(COUNT(*), 0) + 1 AS Count
            #     FROM CommonHeaderTbl
            #     WHERE PermitId LIKE %s
            #     """,
            #     [f"{Username}{refDate}%"]
            # )
           
           
           
            # RefId = "%03d" % (ref_rows[0]['Count'] if ref_rows else 1)

           
           
            # RefId = "%03d" % (ref_rows[0]['Count'] if ref_rows else 1)
# End

            count_rows = SqlDb.execute_query(
                """
                SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                FROM CommonHeaderTbl
                WHERE PermitId LIKE %s
                """,
                [f"{Username}{refDate}%"]
            )
            count = count_rows[0]['Count'] if count_rows else 1

            # 3. All 3 IDs from same count
            RefId    = f"{count:03d}"
            PermitId = f"{Username}{refDate}{RefId}"
            JobId    = f"K{yy_mmdd}{count:05d}"
            MsgId    = f"{refDate}{count:04d}"

            print("PermitId:", PermitId, "| JobId:", JobId, "| MsgId:", MsgId, "| count:", count)


            # job_rows = SqlDb.execute_query(
            #     """
            #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
            #     FROM PermitCount 
            #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'OUTDEC'
            #     """,
            #     [f"{jobDate}%", AccountId] 
            # )
            # JobIdCount = job_rows[0]['Count'] if job_rows else 1
            # JobId = f"K{datetime.now().strftime('%y%m%d')}{JobIdCount:05d}"
            
            # mailbox_rows = SqlDb.execute_query(
            #     "SELECT MailBoxId FROM ManageUser WHERE UserName = %s",
            #     [Username]
            # )
            # MailBoxId = mailbox_rows[0]['MailBoxId'] if mailbox_rows else ""
            # msg_rows = SqlDb.execute_query(
            #     """
            #     SELECT ISNULL(MAX(CAST(RIGHT(MsgId, 4) AS INT)), 0) + 1 AS Count
            #     FROM PermitCount
            #     WHERE AccountId = %s
            #     """,
            #     [AccountId]
            # )
            # MsgCount = msg_rows[0]['Count'] if msg_rows else 1

            # MsgId = f"{datetime.now().strftime('%Y%m%d')}{MsgCount:04d}"

            # PermitId = f"{Username}{refDate}{RefId}"

            # print("PermitId:", PermitId)
            # print("JobId:", JobId)
            # print("MsgId:", MsgId)
            # print("RefId:", RefId)
            # print('AccountId:', AccountId)

            query_join = """
                SELECT TOP 1 
                    manageuser.LoginStatus, manageuser.DateLastUpdated, manageuser.MailBoxId, 
                    manageuser.SeqPool, SequencePool.StartSequence, DeclarantCompany.TradeNetMailboxID, 
                    DeclarantCompany.DeclarantName, DeclarantCompany.DeclarantCode, 
                    DeclarantCompany.DeclarantTel, DeclarantCompany.CRUEI, DeclarantCompany.Code, 
                    DeclarantCompany.name, DeclarantCompany.name1 
                FROM manageuser 
                INNER JOIN SequencePool ON manageuser.SeqPool = SequencePool.Description 
                INNER JOIN DeclarantCompany ON DeclarantCompany.TradeNetMailboxID = ManageUser.MailBoxId 
                WHERE ManageUser.UserName = %s
            """
            head_rows = SqlDb.execute_query(query_join, [Username])

            if not head_rows:
                return Response({"error": "Company profile data not found"}, status=404)

            head = head_rows[0]

            return Response({
                "UserName": Username,
                "PermitId": PermitId,
                "JobId": JobId,
                "RefId": RefId,
                "MsgId": MsgId,
                "AccountId": AccountId,
                "LoginStatus": head.get("LoginStatus", ""),
                "DateLastUpdated": str(head.get("DateLastUpdated", "")),
                "MailBoxId": head.get("MailBoxId", ""),
                "SeqPool": head.get("SeqPool", ""),
                "StartSequence": head.get("StartSequence", ""),
                "TradeNetMailboxID": head.get("TradeNetMailboxID", ""),
                "DeclarantName": head.get("DeclarantName", ""),
                "DeclarantCode": head.get("DeclarantCode", ""),
                "DeclarantTel": head.get("DeclarantTel", ""),
                "CRUEI": head.get("CRUEI", ""),
                "Code": head.get("Code", ""),
                "name": head.get("name", ""),
                "name1": head.get("name1", ""),
                "PermitNumber": "",
                "prmtStatus": "NEW",
                "CurrentDate": currentDate
            })

        except Exception as e:
            print("--- DATABASE/LOGIC ERROR ---")
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)




# class CopyOut(APIView):
#     def post(self, request):
#         try:
#             permits = request.data.get("permits", [])
#             username = request.data.get("user")
#             touch_time = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#             if not permits:
#                 return Response({"error": "No permits selected"}, status=400)
#             if not username:
#                 return Response({"error": "User required"}, status=400)
#             now = datetime.now()
#             ref_date   = now.strftime("%Y%m%d")
#             job_date   = now.strftime("%y%m%d")
#             today_dash = now.strftime("%Y-%m-%d")
#             copied_permits = []
#             with transaction.atomic():
#                 cursor = connection.cursor()
#                 cursor.execute("""
#                     SELECT AccountId, MailBoxId
#                     FROM ManageUser
#                     WHERE UserName = %s
#                 """, [username])
#                 row = cursor.fetchone()
#                 if not row:
#                     return Response({"error": f"User '{username}' not found"}, status=404)
#                 account_id, mailbox_id = row
#                 for permit_id in permits:
#                     cursor.execute("""
#                         SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s
#                     """, [permit_id])
#                     if not cursor.fetchone():
#                         continue
#                     # RefId — count INPDEC records for today
#                     # cursor.execute("""
#                     #     SELECT ISNULL(COUNT(*), 0) + 1
#                     #     FROM CommonHeaderTbl
#                     #     WHERE MSGId LIKE %s AND MessageType = 'INPDEC'
#                     # """, [f"%{ref_date}%"])

#                     cursor.execute("""
#                         SELECT ISNULL(COUNT(*), 0) + 1 AS NextSeq
#                         FROM CommonHeaderTbl
#                         WHERE PermitId LIKE %s
#                     """, [f"{username}{ref_date}%"])
#                     # seq = cursor.fetchone()[0]
#                     count = cursor.fetchone()[0]

#                     # ref_count = cursor.fetchone()[0]

#                     # ref_id = f"{ref_count:03d}"
#                     # JobId + MsgId — scoped to AccountId + today in PermitCount
#                   # 6-6-26 start
#                     # cursor.execute("""
#                     #     SELECT ISNULL(COUNT(*), 0) + 1
#                     #     FROM PermitCount
#                     #     WHERE TouchTime LIKE %s AND AccountId = %s
#                     # """, [f"%{today_dash}%", account_id])
#                     # job_count = cursor.fetchone()[0]
#                     # job_id        = f"K{job_date}{job_count:05d}"
#                     # msg_id        = f"{ref_date}{job_count:04d}"
#                     ref_id        = f"{count:03d}"
#                     job_id        = f"K{job_date}{count:05d}"
#                     msg_id        = f"{ref_date}{count:04d}"
#                     new_permit_id = f"{username}{ref_date}{ref_id}"
#                     cursor.execute("""
#                         INSERT INTO CommonHeaderTbl (
#                             Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
#                             MessageType, DeclarationType, PreviousPermit, CargoPackType,
#                             InwardTransportMode, OutwardTransportMode, BGIndicator,
#                             SupplyIndicator, ReferenceDocuments, License,
#                             COType, Entryyear, GSPDonorCountry,
#                             CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
#                             PerCommon, CurrencyCode, AddCerDtl, TransDtl,
#                             Recipient, DeclarantCompanyCode, ExporterCompanyCode,
#                             Inwardcarriercode, OutwardCarrierAgentCode,
#                             FreightForwarderCode, ImporterCompanyCode,InwardCarrierAgentCode,
#                             CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
#                             HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
#                             VoyageNumber, VesselName, OceanBillofLadingNo,
#                             ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
#                             MasterAirwayBill, ReleaseLocation, RecepitLocation,
#                             StorageLocation, BlanketStartDate,
#                             DepartureDate, DepartureTime, DischargePort,
#                             FinalDestinationCountry, OutVoyageNumber, OutVesselName,
#                             OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
#                             VesselNationality, TowingVesselID, TowingVesselName,
#                             NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
#                             OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
#                             TotalOuterPack, TotalOuterPackUOM,
#                             TotalGrossWeight, TotalGrossWeightUOM,
#                             GrossReference, TradeRemarks, InternalRemarks,
#                             DeclareIndicator, NumberOfItems,
#                             TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
#                             TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
#                             Status, TouchUser, TouchTime,
#                             PermitNumber, prmtStatus,
#                             ResLoaName, RepLocName, RecepitLocName,
#                             outHAWB, INHAWB, seastore, CertificateNumber,
#                             Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
#                             CondColor, TransmitId, gstVerified
#                         )
#                         SELECT
#                             %s, %s, %s, %s, %s,
#                             MessageType, DeclarationType, PreviousPermit, CargoPackType,
#                             InwardTransportMode, OutwardTransportMode, BGIndicator,
#                             SupplyIndicator, ReferenceDocuments, License,
#                             COType, Entryyear, GSPDonorCountry,
#                             CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
#                             PerCommon, CurrencyCode, AddCerDtl, TransDtl,
#                             Recipient, DeclarantCompanyCode, ExporterCompanyCode,
#                             NULL, OutwardCarrierAgentCode,
#                             FreightForwarderCode, ImporterCompanyCode,InwardCarrierAgentCode,
#                             CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
#                             HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
#                             VoyageNumber, VesselName, OceanBillofLadingNo,
#                             ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
#                             MasterAirwayBill, ReleaseLocation, RecepitLocation,
#                             StorageLocation, BlanketStartDate,
#                             DepartureDate, DepartureTime, DischargePort,
#                             FinalDestinationCountry, OutVoyageNumber, OutVesselName,
#                             OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
#                             VesselNationality, TowingVesselID, TowingVesselName,
#                             NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
#                             OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
#                             TotalOuterPack, TotalOuterPackUOM,
#                             TotalGrossWeight, TotalGrossWeightUOM,
#                             GrossReference, TradeRemarks, InternalRemarks,
#                             DeclareIndicator, NumberOfItems,
#                             TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
#                             TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
#                             'DRF', %s, %s,
#                             NULL, 'NEW',
#                             ResLoaName, RepLocName, RecepitLocName,
#                             outHAWB, INHAWB, seastore, CertificateNumber,
#                             Defrentprinting, Cnb, NULL, MRDate, MRTime,
#                             CondColor, TransmitId, gstVerified
#                         FROM CommonHeaderTbl
#                         WHERE PermitId = %s
#                     """, [
#                         ref_id, job_id, msg_id, new_permit_id, mailbox_id,
#                         username, touch_time,
#                         permit_id
#                     ])
#                     child_tables = {
#                         "CommonInvoiceDtl": [
#                             "SNo", "InvoiceNo", "InvoiceDate", "TermType",
#                             "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
#                             "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
#                             "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
#                             "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
#                             "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
#                             "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
#                             "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
#                         ],
#                         "CommonItemDtl": [
#                             "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
#                             "Contry", "EndUserDescription", "Brand", "Model",
#                             "InHAWBOBL", "OutHAWBOBL", "DutiableQty", "DutiableUOM",
#                             "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
#                             "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
#                             "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
#                             "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
#                             "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM", "ImPQty", "ImPUOM",
#                             "PreferentialCode", "GSTRate", "GSTUOM", "GSTAmount",
#                             "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
#                             "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
#                             "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
#                             "CurrentLot", "PreviousLot", "LSPValue", "Making",
#                             "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
#                             "CerItemQty", "CerItemUOM", "CIFValOfCer", "ManufactureCostDate",
#                             "TexCat", "TexQuotaQty", "TexQuotaUOM", "CerInvNo", "CerInvDate",
#                             "OriginOfCer", "HSCodeCer", "PerContent", "CertificateDescription",
#                             "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
#                             "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
#                             "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
#                         ],
#                         "CommonCASCDtl": [
#                                 "ItemNo", "ProductCode", "Quantity", "ProductUOM",
#                                 "RowNo", "CascCode1", "CascCode2", "CascCode3",
#                                 "MessageType", "TouchUser", "TouchTime", "EndUserDes", "CASCId"
#                             ],
#                         "CommonCPCDtl": [
#                             "MessageType", "RowNo", "CPCType",
#                             "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
#                             "TouchUser", "TouchTime"
#                         ],
#                         "CommonContainerDtl": [
#                             "RowNo", "ContainerNo", "Size", "Weight", "SealNo", "MessageType","TouchUser", "TouchTime"
#                         ],
#                         "CommonFile": [
#                             "Name", "ContentType", "Data", "DocumentType",
#                             "TouchUser", "TouchTime", "filePath", "Size", "Type"
#                         ],
#                         "CommonPMT": [
#                             "ConditionCode", "ConditionDesc", "PermitNumber",
#                         ],
#                     }

#                     for table, cols in child_tables.items():
#                         col_list    = ", ".join(["PermitId"] + cols)
#                         select_cols = ", ".join(cols)
#                         try:
#                             cursor.execute(f"""
#                                 INSERT INTO {table} ({col_list})
#                                 SELECT %s, {select_cols}
#                                 FROM {table}
#                                 WHERE PermitId = %s
#                             """, [new_permit_id, permit_id])
#                         except Exception as err:
#                             print(f"Warning copying {table}: {err}")

#                     cursor.execute("""
#                         INSERT INTO PermitCount
#                             (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
#                         VALUES (%s, 'OUTDEC', %s, %s, %s, %s)
#                     """, [new_permit_id, account_id, msg_id, username, touch_time])

#                     copied_permits.append(new_permit_id)

#             return Response({
#                 "SUCCESS": True,
#                 "message": f"{len(copied_permits)} permit(s) copied successfully",
#                 "copiedPermits": copied_permits,
#             })

#         except Exception as e:
#             import traceback
#             traceback.print_exc()
#         return Response({
#             "SUCCESS": True,
#             "message": f"{len(copied_permits)} permit(s) copied successfully",
#             "copiedPermits": copied_permits,
#         })


# class CopyOut(APIView):
#     def post(self, request):
#         copied_permits = []
#         try:
#             permits = request.data.get("permits", [])
#             username = request.data.get("user")
#             touch_time = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#             if not permits:
#                 return Response({"error": "No permits selected"}, status=400)
#             if not username:
#                 return Response({"error": "User required"}, status=400)
#             now = datetime.now()
#             ref_date   = now.strftime("%Y%m%d")
#             job_date   = now.strftime("%y%m%d")
#             today_dash = now.strftime("%Y-%m-%d")

#             with transaction.atomic():
#                 cursor = connection.cursor()
#                 cursor.execute("""
#                     SELECT AccountId, MailBoxId
#                     FROM ManageUser
#                     WHERE UserName = %s
#                 """, [username])
#                 row = cursor.fetchone()
#                 if not row:
#                     return Response({"error": f"User '{username}' not found"}, status=404)
#                 account_id, mailbox_id = row

#                 # GLOBAL starting count for JobId/MsgId
#                 cursor.execute("""
#                     SELECT ISNULL(COUNT(*), 0) + 1 AS Count
#                     FROM CommonHeaderTbl
#                     WHERE JobId LIKE %s
#                 """, [f"K{job_date}%"])
#                 job_count = cursor.fetchone()[0]

#                 # Per-user starting count for PermitId/RefId
#                 cursor.execute("""
#                     SELECT ISNULL(COUNT(*), 0) + 1 AS NextSeq
#                     FROM CommonHeaderTbl
#                     WHERE PermitId LIKE %s
#                 """, [f"{username}{ref_date}%"])
#                 ref_count = cursor.fetchone()[0]

#                 for permit_id in permits:
#                     cursor.execute("""
#                         SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s
#                     """, [permit_id])
#                     if not cursor.fetchone():
#                         continue

#                     ref_id        = f"{ref_count:03d}"
#                     job_id        = f"K{job_date}{job_count:05d}"
#                     msg_id        = f"{ref_date}{job_count:04d}"
#                     new_permit_id = f"{username}{ref_date}{ref_id}"

#                     cursor.execute("""
#                         INSERT INTO CommonHeaderTbl (
#                             Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
#                             MessageType, DeclarationType, PreviousPermit, CargoPackType,
#                             InwardTransportMode, OutwardTransportMode, BGIndicator,
#                             SupplyIndicator, ReferenceDocuments, License,
#                             COType, Entryyear, GSPDonorCountry,
#                             CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
#                             PerCommon, CurrencyCode, AddCerDtl, TransDtl,
#                             Recipient, DeclarantCompanyCode, ExporterCompanyCode,
#                             Inwardcarriercode, OutwardCarrierAgentCode,
#                             FreightForwarderCode, ImporterCompanyCode,InwardCarrierAgentCode,
#                             CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
#                             HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
#                             VoyageNumber, VesselName, OceanBillofLadingNo,
#                             ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
#                             MasterAirwayBill, ReleaseLocation, RecepitLocation,
#                             StorageLocation, BlanketStartDate,
#                             DepartureDate, DepartureTime, DischargePort,
#                             FinalDestinationCountry, OutVoyageNumber, OutVesselName,
#                             OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
#                             VesselNationality, TowingVesselID, TowingVesselName,
#                             NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
#                             OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
#                             TotalOuterPack, TotalOuterPackUOM,
#                             TotalGrossWeight, TotalGrossWeightUOM,
#                             GrossReference, TradeRemarks, InternalRemarks,
#                             DeclareIndicator, NumberOfItems,
#                             TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
#                             TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
#                             Status, TouchUser, TouchTime,
#                             PermitNumber, prmtStatus,
#                             ResLoaName, RepLocName, RecepitLocName,
#                             outHAWB, INHAWB, seastore, CertificateNumber,
#                             Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
#                             CondColor, TransmitId, gstVerified
#                         )
#                         SELECT
#                             %s, %s, %s, %s, %s,
#                             MessageType, DeclarationType, PreviousPermit, CargoPackType,
#                             InwardTransportMode, OutwardTransportMode, BGIndicator,
#                             SupplyIndicator, ReferenceDocuments, License,
#                             COType, Entryyear, GSPDonorCountry,
#                             CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
#                             PerCommon, CurrencyCode, AddCerDtl, TransDtl,
#                             Recipient, DeclarantCompanyCode, ExporterCompanyCode,
#                             NULL, OutwardCarrierAgentCode,
#                             FreightForwarderCode, ImporterCompanyCode,InwardCarrierAgentCode,
#                             CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
#                             HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
#                             VoyageNumber, VesselName, OceanBillofLadingNo,
#                             ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
#                             MasterAirwayBill, ReleaseLocation, RecepitLocation,
#                             StorageLocation, BlanketStartDate,
#                             DepartureDate, DepartureTime, DischargePort,
#                             FinalDestinationCountry, OutVoyageNumber, OutVesselName,
#                             OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
#                             VesselNationality, TowingVesselID, TowingVesselName,
#                             NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
#                             OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
#                             TotalOuterPack, TotalOuterPackUOM,
#                             TotalGrossWeight, TotalGrossWeightUOM,
#                             GrossReference, TradeRemarks, InternalRemarks,
#                             DeclareIndicator, NumberOfItems,
#                             TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
#                             TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
#                             'DRF', %s, %s,
#                             NULL, 'NEW',
#                             ResLoaName, RepLocName, RecepitLocName,
#                             outHAWB, INHAWB, seastore, CertificateNumber,
#                             Defrentprinting, Cnb, NULL, MRDate, MRTime,
#                             CondColor, TransmitId, gstVerified
#                         FROM CommonHeaderTbl
#                         WHERE PermitId = %s
#                     """, [
#                         ref_id, job_id, msg_id, new_permit_id, mailbox_id,
#                         username, touch_time,
#                         permit_id
#                     ])

#                     child_tables = {
#                         "CommonInvoiceDtl": [
#                             "SNo", "InvoiceNo", "InvoiceDate", "TermType",
#                             "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
#                             "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
#                             "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
#                             "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
#                             "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
#                             "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
#                             "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
#                         ],
#                         "CommonItemDtl": [
#                             "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
#                             "Contry", "EndUserDescription", "Brand", "Model",
#                             "InHAWBOBL", "OutHAWBOBL", "DutiableQty", "DutiableUOM",
#                             "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
#                             "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
#                             "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
#                             "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
#                             "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM", "ImPQty", "ImPUOM",
#                             "PreferentialCode", "GSTRate", "GSTUOM", "GSTAmount",
#                             "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
#                             "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
#                             "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
#                             "CurrentLot", "PreviousLot", "LSPValue", "Making",
#                             "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
#                             "CerItemQty", "CerItemUOM", "CIFValOfCer", "ManufactureCostDate",
#                             "TexCat", "TexQuotaQty", "TexQuotaUOM", "CerInvNo", "CerInvDate",
#                             "OriginOfCer", "HSCodeCer", "PerContent", "CertificateDescription",
#                             "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
#                             "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
#                             "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
#                         ],
#                         "CommonCASCDtl": [
#                                 "ItemNo", "ProductCode", "Quantity", "ProductUOM",
#                                 "RowNo", "CascCode1", "CascCode2", "CascCode3",
#                                 "MessageType", "TouchUser", "TouchTime", "EndUserDes", "CASCId"
#                             ],
#                         "CommonCPCDtl": [
#                             "MessageType", "RowNo", "CPCType",
#                             "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
#                             "TouchUser", "TouchTime"
#                         ],
#                         "CommonContainerDtl": [
#                             "RowNo", "ContainerNo", "Size", "Weight", "SealNo", "MessageType","TouchUser", "TouchTime"
#                         ],
#                         "CommonFile": [
#                             "Name", "ContentType", "Data", "DocumentType",
#                             "TouchUser", "TouchTime", "filePath", "Size", "Type"
#                         ],
#                         "CommonPMT": [
#                             "ConditionCode", "ConditionDesc", "PermitNumber",
#                         ],
#                     }

#                     for table, cols in child_tables.items():
#                         col_list    = ", ".join(["PermitId"] + cols)
#                         select_cols = ", ".join(cols)
#                         try:
#                             cursor.execute(f"""
#                                 INSERT INTO {table} ({col_list})
#                                 SELECT %s, {select_cols}
#                                 FROM {table}
#                                 WHERE PermitId = %s
#                             """, [new_permit_id, permit_id])
#                         except Exception as err:
#                             print(f"Warning copying {table}: {err}")

#                     cursor.execute("""
#                         INSERT INTO PermitCount
#                             (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
#                         VALUES (%s, 'OUTDEC', %s, %s, %s, %s)
#                     """, [new_permit_id, account_id, msg_id, username, touch_time])

#                     copied_permits.append(new_permit_id)

#                     job_count += 1
#                     ref_count += 1

#             return Response({
#                 "SUCCESS": True,
#                 "message": f"{len(copied_permits)} permit(s) copied successfully",
#                 "copiedPermits": copied_permits,
#             })

#         except Exception as e:
#             import traceback
#             traceback.print_exc()
#             return Response({"error": f"Database Error: {str(e)}"}, status=400)

class CopyOut(APIView):
    def post(self, request):
        copied_permits = []
        try:
            permits = request.data.get("permits", [])
            username = request.data.get("user")
            touch_time = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if not permits:
                return Response({"error": "No permits selected"}, status=400)
            if not username:
                return Response({"error": "User required"}, status=400)
            now = datetime.now()
            ref_date = now.strftime("%Y%m%d")
            job_date = now.strftime("%y%m%d")

            with transaction.atomic():
                cursor = connection.cursor()
                cursor.execute("""
                    SELECT AccountId, MailBoxId
                    FROM ManageUser
                    WHERE UserName = %s
                """, [username])
                row = cursor.fetchone()
                if not row:
                    return Response({"error": f"User '{username}' not found"}, status=404)
                account_id, mailbox_id = row

                # GLOBAL starting count for JobId/MsgId
                cursor.execute("""
                    SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                    FROM CommonHeaderTbl
                    WHERE JobId LIKE %s
                """, [f"K{job_date}%"])
                job_count = cursor.fetchone()[0]

                # Per-user starting count for PermitId/RefId
                cursor.execute("""
                    SELECT ISNULL(COUNT(*), 0) + 1 AS NextSeq
                    FROM CommonHeaderTbl
                    WHERE PermitId LIKE %s
                """, [f"{username}{ref_date}%"])
                ref_count = cursor.fetchone()[0]

                for permit_id in permits:
                    cursor.execute("""
                        SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s
                    """, [permit_id])
                    if not cursor.fetchone():
                        continue

                    ref_id        = f"{ref_count:03d}"
                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    # ── 1. Insert new CommonHeaderTbl row (unchanged logic) ──
                    cursor.execute("""
                        INSERT INTO CommonHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode,InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        )
                        SELECT
                            %s, %s, %s, %s, %s,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            NULL, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode,InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            'DRF', %s, %s,
                            NULL, 'NEW',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, NULL, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. NEW: mirror the row we just inserted into OutHeaderTbl ──
                    # Column list per OutHeaderTbl (note: no HBL, ClaimantPartyCode,
                    # seastore, gstVerified, Message, HandlingAgentCode, CustomerRemarks).
                    cursor.execute("""
                        INSERT INTO OutHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, EndUserCode, Manufacturer,
                            ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, EndUserCode, Manufacturer,
                            ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: copy into BOTH Common* and Out* targets ──
                    # Each entry: source_table -> (out_target_table, cols).
                    # cols are the CommonXxx column names; for tables whose Out*
                    # counterpart uses different column names (Invoice/File) we
                    # remap explicitly below instead of using this generic loop.
                    child_tables = {
                        "CommonItemDtl": ("OutItemDtl", [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "EndUserDescription", "Brand", "Model",
                            "InHAWBOBL", "OutHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode",
                            "GSTRate", "GSTUOM", "GSTAmount",
                            "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "CerItemQty", "CerItemUOM", "CIFValOfCer", "ManufactureCostDate",
                            "TexCat", "TexQuotaQty", "TexQuotaUOM", "CerInvNo", "CerInvDate",
                            "OriginOfCer", "HSCodeCer", "PerContent", "CertificateDescription",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                            # NOTE: LSPValue exists in CommonItemDtl but has no
                            # counterpart column in OutItemDtl, so it's dropped here.
                        ]),
                        "CommonCASCDtl": ("OutCASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
                            "RowNo", "CascCode1", "CascCode2", "CascCode3",
                            "MessageType", "TouchUser", "TouchTime", "CASCId", "EndUserDes"
                        ]),
                        "CommonCPCDtl": ("OutCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("OutContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, (dst_table, cols) in child_tables.items():
                        col_list = ", ".join(cols)
                        # Common -> Common (from OLD permit, as in the original logic)
                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")

                        # Common(NEW) -> Out target — reads what we just wrote,
                        # so both tables end up identical for this permit.
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── 4. Invoice — column name differs (ExportPartyCode vs
                    # ImportPartyCode) and OutInvoiceDtl has no ChkOtherInv ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonInvoiceDtl (
                                PermitId, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime, ChkOtherInv
                            )
                            SELECT
                                %s, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime, ChkOtherInv
                            FROM CommonInvoiceDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonInvoiceDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO OutInvoiceDtl (
                                PermitId, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ExportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime
                            )
                            SELECT
                                %s, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime
                            FROM CommonInvoiceDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonInvoiceDtl -> OutInvoiceDtl: {err}")

                    # ── 5. File — OutFile has no filePath, and uses InPaymentId
                    # instead of PaymentId ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonFile: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO OutFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonFile -> OutFile: {err}")

                    # ── 6. PMT — copy Common only (no OutPMT columns supplied) ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonPMT (PermitId, ConditionCode, ConditionDesc, PermitNumber)
                            SELECT %s, ConditionCode, ConditionDesc, PermitNumber
                            FROM CommonPMT WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying PMT: {err}")

                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'OUTDEC', %s, %s, %s, %s)
                    """, [new_permit_id, account_id, msg_id, username, touch_time])

                    copied_permits.append(new_permit_id)
                    job_count += 1
                    ref_count += 1

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) copied successfully",
                "copiedPermits": copied_permits,
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": f"Database Error: {str(e)}"}, status=400)



# ── Page dimensions ──────────────────────────────────────────────────────────
W, H = 595, 842   # A4 in points

# ── Exact column X positions (derived from pdfplumber line analysis) ─────────
L_BORDER = 23
R_BORDER = 572
MID_DIV  = 314    # vertical divider separating left boxes from right (title / box4 / box12)

COL_X_PDF = [23, 73, 138, 314, 385, 489, 572]
#             │    │    │    │    │    │    └─ right border
#             │    │    │    │    │    └─ invoice col left edge
#             │    │    │    │    └─ gross-weight col left edge
#             │    │    │    └─ origin-criterion col left edge
#             │    │    └─ description col left edge
#             │    └─ marks col left edge
#             └─ left border / item-no col left edge


def _rl(pdf_top: float) -> float:
    """Convert pdfplumber top-origin coord → ReportLab bottom-origin y."""
    return H - pdf_top


# ════════════════════════════════════════════════════════════════════════════
# DJANGO VIEW
# ════════════════════════════════════════════════════════════════════════════

class DraftCoo(APIView):

    def get(self, request, PermitId):
        try:
            # ── 1. CommonHeaderTbl ────────────────────────────────────────────
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT
                        JobId, MSGId, PermitNumber, TouchTime,
                        COType,
                        ExporterCompanyCode, ImporterCompanyCode, CONSIGNEECode,
                        VesselName, FlightNO,
                        InwardTransportMode, OutwardTransportMode,
                        DepartureDate, DepartureTime,
                        DischargePort,
                        FinalDestinationCountry,
                        LoadingPortCode,
                        GSPDonorCountry,
                        DeclarantCompanyCode,
                        CertificateNumber,
                        DeclarationType
                    FROM CommonHeaderTbl
                    WHERE PermitId = %s
                """, [PermitId])
                hrow = cursor.fetchone()

            if not hrow:
                return HttpResponse("Permit not found", status=404)

            (
                JobId, MSGId, PermitNumber, TouchTime,
                COType,
                ExporterCode, ImporterCode, ConsigneeCode,
                VesselName, FlightNO,
                InwardTransportMode, OutwardTransportMode,
                DepartureDate, DepartureTime,
                DischargePort,
                FinalDestinationCountry,
                LoadingPortCode,
                GSPDonorCountry,
                DeclarantCode,
                CertificateNumber,
                DeclarationType,
            ) = hrow

            # ── 2. CommonItemDtl ──────────────────────────────────────────────
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT
                        ItemNo, HSCode, Description, CertificateDescription,
                        ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                        InvoiceNo, CerInvNo, CerInvDate,
                        OriginOfCer, HSCodeCer, PreferentialCode,
                        CerItemQty, CerItemUOM,
                        CIFValOfCer,
                        DutiableQty, DutiableUOM,
                        InvoiceQuantity, Contry
                    FROM CommonItemDtl
                    WHERE PermitId = %s
                    ORDER BY ItemNo
                """, [PermitId])
                item_rows = cursor.fetchall()

            if not item_rows:
                return HttpResponse("No item details found", status=404)

            # ── 3. CommonInvoiceDtl ───────────────────────────────────────────
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT
                        InvoiceNo, InvoiceDate, TermType, ImportPartyCode,
                        TICurrency, TIExRate, TIAmount, TISAmount,
                        FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                        ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                        OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                        CIFSUMAmount, GSTPercentage, GSTSUMAmount
                    FROM CommonInvoiceDtl
                    WHERE PermitId = %s
                """, [PermitId])
                inv_rows = cursor.fetchall()

            # ── 4. Derive transport & dates ───────────────────────────────────
            mode_map   = {"1": "SEA", "2": "AIR", "3": "LAND", "4": "RAIL", "5": "POST"}
            trans_mode = mode_map.get(self._sv(OutwardTransportMode),
                                      self._sv(OutwardTransportMode))
            vessel = self._sv(VesselName) or self._sv(FlightNO)

            dep_date = ""
            if DepartureDate:
                try:    dep_date = DepartureDate.strftime("%d/%m/%Y")
                except: dep_date = self._sv(DepartureDate)

            touch_str = ""
            if TouchTime:
                try:    touch_str = TouchTime.strftime("%d/%m/%Y %H:%M:%S")
                except: touch_str = self._sv(TouchTime)

            # ── 5. COType → Form label ────────────────────────────────────────
            form_map = {
                "D":  ("FORM D",  "ASEAN TRADE IN GOODS AGREEMENT /\nASEAN INDUSTRIAL COOPERATION"),
                "E":  ("FORM E",  "ASEAN-CHINA FREE TRADE AREA\nPREFERENTIAL TARIFF"),
                "AI": ("FORM AI", "ASEAN-INDIA FREE TRADE AREA\nPREFERENTIAL TARIFF"),
                "AK": ("FORM AK", "ASEAN-KOREA FREE TRADE AREA\nPREFERENTIAL TARIFF"),
            }
            co_key = self._sv(COType).upper()
            form_type, agreement = form_map.get(
                co_key, (f"FORM {co_key}", "ASEAN TRADE IN GOODS AGREEMENT"))

            # ── 6. Invoice lookup ─────────────────────────────────────────────
            inv_lookup: dict = {}
            for inv in (inv_rows or []):
                key = self._sv(inv[0])
                if key and key not in inv_lookup:
                    inv_lookup[key] = inv

            # ── 7. Build items list ───────────────────────────────────────────
            items = []
            for row in item_rows:
                (
                    ItemNo, HSCode, Description, CertDesc,
                    Marks1, Marks2, Marks3, Marks4,
                    InvoiceNo, CerInvNo, CerInvDate,
                    OriginOfCer, HSCodeCer, PreferentialCode,
                    CerItemQty, CerItemUOM,
                    CIFValOfCer,
                    DutiableQty, DutiableUOM,
                    InvoiceQuantity, Contry,
                ) = row

                marks = "\n".join(
                    m for m in [
                        self._sv(Marks1), self._sv(Marks2),
                        self._sv(Marks3), self._sv(Marks4),
                    ] if m
                ) or "N/M"

                inv_ref  = inv_lookup.get(self._sv(CerInvNo) or self._sv(InvoiceNo), {})
                cif_val  = self._sv(CIFValOfCer)
                currency = self._sv(inv_ref[4]) if inv_ref else ""

                gross_weight = (
                    f"{self._sv(CerItemQty)} {self._sv(CerItemUOM)}"
                    if self._sv(CerItemQty) else
                    f"{self._sv(DutiableQty)} {self._sv(DutiableUOM)}"
                )
                fob_value = f"{cif_val} {currency}".strip() if cif_val else ""

                inv_date_str = ""
                if CerInvDate:
                    try:    inv_date_str = CerInvDate.strftime("%d/%m/%Y")
                    except: inv_date_str = self._sv(CerInvDate)

                items.append({
                    "item_no":          self._sv(ItemNo),
                    "marks":            marks,
                    "description":      self._sv(CertDesc) or self._sv(Description),
                    "hs_code":          self._sv(HSCodeCer) or self._sv(HSCode),
                    "net_weight":       f"{self._sv(CerItemQty)} {self._sv(CerItemUOM)}".strip(),
                    "total_line":       "",   # populate if your DB has a total field
                    "extra_lines":      [],   # populate for extra consignee notes etc.
                    "origin_criterion": self._sv(PreferentialCode),
                    "gross_weight":     gross_weight,
                    "fob_value":        fob_value,
                    "invoice_no":       self._sv(CerInvNo) or self._sv(InvoiceNo),
                    "invoice_date":     inv_date_str,
                    "origin_country":   self._sv(OriginOfCer) or self._sv(Contry),
                })

            # ── 8. Assemble data dict ─────────────────────────────────────────
            data = {
                "header_data": {
                    "form_type":         form_type,
                    "agreement":         agreement,
                    "issued_in":         "Singapore",
                    "job_no":            self._sv(JobId),
                    "msg_id":            self._sv(MSGId),
                    "permit_number":     self._sv(PermitNumber),
                    "certificate_no":    self._sv(CertificateNumber),
                    "touch_time":        touch_str,
                    "exporter_name":     self._sv(ExporterCode),   # swap with name lookup
                    "consignee":         self._sv(ConsigneeCode),  # swap with name lookup
                    "transport_mode":    trans_mode,
                    "departure_date":    dep_date,
                    "vessel":            vessel,
                    "port_of_discharge": self._sv(DischargePort),
                },
                "items": items,
                "declaration": {
                    "origin_country":      self._sv(GSPDonorCountry) or "SINGAPORE",
                    "destination_country": self._sv(FinalDestinationCountry),
                    "date":                dep_date,
                },
            }

            # ── 9. Generate PDF ───────────────────────────────────────────────
            pdf_bytes = self._build_pdf(data)
            response  = HttpResponse(pdf_bytes, content_type="application/pdf")
            response["Content-Disposition"] = (
                f'attachment; filename="{self._sv(PermitNumber)}_COO_DRAFT.pdf"'
            )
            return response

        except Exception as e:
            import traceback
            traceback.print_exc()
            return HttpResponse(f"Error generating COO PDF: {str(e)}", status=500)

    # ══════════════════════════════════════════════════════════════════════════
    # OPTIONAL: resolve company code → name + address
    # ══════════════════════════════════════════════════════════════════════════
    # def _get_party_name(self, code):
    #     if not code:
    #         return ""
    #     with connection.cursor() as cursor:
    #         cursor.execute(
    #             "SELECT Name, Address FROM CommonPartyTbl WHERE Code = %s", [code])
    #         row = cursor.fetchone()
    #     if row:
    #         return "\n".join(filter(None, [str(row[0] or ""), str(row[1] or "")]))
    #     return code

    # ══════════════════════════════════════════════════════════════════════════
    # PDF BUILDER
    # ══════════════════════════════════════════════════════════════════════════

    def _build_pdf(self, data: dict) -> bytes:
        packet = io.BytesIO()
        can    = canvas.Canvas(packet, pagesize=(W, H))

        self._draw_outer_border(can)
        self._draw_draft_watermark(can)
        self._draw_draft_notices(can)

        self._draw_top_banner(can, data)
        self._draw_box2(can, data)
        self._draw_box_3_4(can, data)
        self._draw_items_header(can)

        items_end_pdf = self._draw_items_rows(can, data.get("items", []))

        BOX_11_TOP = 521
        if items_end_pdf < BOX_11_TOP:
            self._draw_filler(can, items_end_pdf, BOX_11_TOP)

        self._draw_boxes_11_12(can, BOX_11_TOP, data)

        can.save()
        packet.seek(0)
        return packet.read()

    # ══════════════════════════════════════════════════════════════════════════
    # DRAW HELPERS
    # ══════════════════════════════════════════════════════════════════════════

    def _sv(self, val):
        if val is None:
            return ""
        if str(val).strip() in ("", "--Select--"):
            return ""
        return str(val).strip()

    def _line(self, can, x1, y1, x2, y2, width=0.5):
        can.setLineWidth(width)
        can.line(x1, y1, x2, y2)

    def _bold(self, can, x, y, text, size=7):
        can.setFont("Helvetica-Bold", size)
        can.drawString(x, y, text)

    def _normal(self, can, x, y, text, size=7.5):
        can.setFont("Helvetica", size)
        can.drawString(x, y, text)

    def _centred(self, can, cx, y, text, font="Helvetica", size=7.5):
        can.setFont(font, size)
        can.drawCentredString(cx, y, str(text))

    # ══════════════════════════════════════════════════════════════════════════
    # DRAW SECTIONS
    # ══════════════════════════════════════════════════════════════════════════

    def _draw_outer_border(self, can):
        can.setLineWidth(0.8)
        can.rect(L_BORDER, _rl(737), R_BORDER - L_BORDER, _rl(28) - _rl(737))

    def _draw_draft_watermark(self, can):
        can.saveState()
        can.setFont("Helvetica-Bold", 80)
        can.setFillColorRGB(0.85, 0.85, 0.85)
        can.setFillAlpha(0.30)
        can.translate(W / 2, H / 2)
        can.rotate(45)
        can.drawCentredString(0, 0, "DRAFT")
        can.restoreState()

    def _draw_draft_notices(self, can):
        notice = "THIS IS A DRAFT CERTIFICATE OF ORIGIN AND IS NOT VALID FOR SUBMISSION."
        can.setFont("Helvetica-Bold", 7)
        can.setFillColorRGB(1, 0, 0)
        can.drawCentredString(W / 2, _rl(749), notice)
        can.setFillColorRGB(0, 0, 0)

    # ── Top banner: Box 1 (left) + Title block (right) ────────────────────────
    # pdf_top band: 28 → 115
    def _draw_top_banner(self, can, data):
        hd        = data.get("header_data", {})
        form_type = self._sv(hd.get("form_type", "FORM E"))
        agreement = self._sv(hd.get("agreement", ""))
        issued_in = self._sv(hd.get("issued_in", "Singapore"))
        job_no    = self._sv(hd.get("job_no", ""))

        top_pdf = 28

        # Vertical divider and bottom rule (left side only, right continues down)
        self._line(can, MID_DIV, _rl(top_pdf), MID_DIV, _rl(115))
        self._line(can, L_BORDER, _rl(115), MID_DIV, _rl(115))

        # ── RIGHT: PG line ────────────────────────────────────────────────────
        can.setFont("Helvetica", 7)
        can.drawRightString(R_BORDER - 4, _rl(top_pdf + 16), "PG : 1 OF 1")

        # ── RIGHT: Job No ─────────────────────────────────────────────────────
        can.setFont("Helvetica", 7)
        can.drawString(MID_DIV + 6, _rl(top_pdf + 16), f"Job No.  {job_no}")

        # ── RIGHT: Agreement + COO title ──────────────────────────────────────
        mid_rx = (MID_DIV + R_BORDER) / 2
        ty_pdf = top_pdf + 51      # first agreement line (≈ pdf_top 79)
        can.setFont("Helvetica-Bold", 8)
        for ln in agreement.split("\n"):
            can.drawCentredString(mid_rx, _rl(ty_pdf), ln.strip())
            ty_pdf += 13
        can.drawCentredString(mid_rx, _rl(ty_pdf), "CERTIFICATE OF ORIGIN")
        ty_pdf += 10
        can.setFont("Helvetica", 7)
        can.drawCentredString(mid_rx, _rl(ty_pdf + 4), "(Combined Declaration and Certificate)")

        # FORM E large grey
        can.setFont("Helvetica-Bold", 6)
        can.setFillColorRGB(0.72, 0.72, 0.72)
        can.drawCentredString(mid_rx, _rl(top_pdf + 113), form_type)
        can.setFillColorRGB(0, 0, 0)

        # "Issued in Singapore"
        can.setFont("Helvetica-Bold", 8)
        can.drawCentredString(mid_rx, _rl(top_pdf + 154), f"Issued in {issued_in}")

        # ── LEFT: Box 1 label + exporter ─────────────────────────────────────
        bx     = L_BORDER + 6
        by_pdf = top_pdf + 8
        self._bold(can, bx, _rl(by_pdf),
                   "1.Goods consigned from (Exporter's business name,address,country)",
                   size=6.5)
        by_pdf += 16
        for ln in self._sv(hd.get("exporter_name", "")).split("\n"):
            if ln:
                self._normal(can, bx + 2, _rl(by_pdf), ln)
                by_pdf += 12

    # ── Box 2: Consignee  (pdf_top 115 → 197) ────────────────────────────────
    def _draw_box2(self, can, data):
        hd = data.get("header_data", {})
        self._line(can, MID_DIV, _rl(115), MID_DIV, _rl(197))
        self._line(can, L_BORDER, _rl(197), R_BORDER, _rl(197))

        bx     = L_BORDER + 6
        by_pdf = 115 + 8
        self._bold(can, bx, _rl(by_pdf),
                   "2.Goods consigned to (Consignee's name, address and country)",
                   size=6.5)
        by_pdf += 16
        for ln in self._sv(hd.get("consignee", "")).split("\n"):
            if ln:
                self._normal(can, bx + 2, _rl(by_pdf), ln)
                by_pdf += 12

    # ── Boxes 3 + 4: Transport / Official Use  (pdf_top 197 → 277) ───────────
    def _draw_box_3_4(self, can, data):
        hd = data.get("header_data", {})
        self._line(can, MID_DIV, _rl(197), MID_DIV, _rl(277))
        self._line(can, L_BORDER, _rl(277), R_BORDER, _rl(277))

        bx = L_BORDER + 6
        self._bold(can, bx, _rl(197 + 8),
                   "3.Means of transport and route (as far as know)", size=6.5)
        self._normal(can, bx + 248, _rl(197 + 8),
                     f"By {self._sv(hd.get('transport_mode', ''))}")

        sub_rows = [
            (213, "Departure Date",           "departure_date",    145),
            (229, "Vessel's Name/Aircraft etc.", "vessel",          147),
            (245, "Port of Discharge",         "port_of_discharge", 145),
        ]
        for pdf_top_row, label, key, val_x in sub_rows:
            self._line(can, L_BORDER, _rl(pdf_top_row), MID_DIV, _rl(pdf_top_row))
            self._line(can, MID_DIV,  _rl(pdf_top_row), R_BORDER, _rl(pdf_top_row))
            self._normal(can, bx + 2, _rl(pdf_top_row + 8), label)
            self._normal(can, val_x,  _rl(pdf_top_row + 8),
                         self._sv(hd.get(key, "")))

        self._bold(can, MID_DIV + 6, _rl(197 + 8), "4.For Official Use", size=6.5)

    # ── Column header  (pdf_top 277 → 329) ───────────────────────────────────
    _COL_DEFS = [
        (23,  73,  ["5.Item", "Number"]),
        (73,  138, ["6.Marks and", "numbers on", "packages"]),
        (138, 314, ["7.Number and type of packages,",
                    "description of goods (inculding",
                    "quantity where appropriate and HS",
                    "number of the importing country)"]),
        (314, 385, ["8.Origin criterion", "(see Overleaf", "Notes)"]),
        (385, 489, ["9.Gross weight or other", "quantity and values",
                    "(FOB) where RVC", "is applied"]),
        (489, 572, ["10.Number and", "date of", "invoices"]),
    ]

    def _draw_items_header(self, can):
        top_pdf, bot_pdf = 277, 329
        for (x1, x2, lines) in self._COL_DEFS:
            self._line(can, x1, _rl(top_pdf), x1, _rl(bot_pdf))
            ty_pdf = top_pdf + 8
            can.setFont("Helvetica-Bold", 6.5)
            for ln in lines:
                can.drawString(x1 + 2, _rl(ty_pdf), ln)
                ty_pdf += 12
        self._line(can, R_BORDER, _rl(top_pdf), R_BORDER, _rl(bot_pdf))
        self._line(can, L_BORDER, _rl(bot_pdf),  R_BORDER, _rl(bot_pdf))

    # ── Item rows  (starts at pdf_top 329) ───────────────────────────────────
    def _draw_items_rows(self, can, items) -> int:
        """Draw all item rows; returns final pdf_top (bottom of last row)."""
        ROW_MIN = 64   # minimum row height in pdf points
        pdf_top = 329

        for item in items:
            # Measure col-2 to determine row height
            col2_w = 314 - 138 - 6
            lines_c2 = []
            hs    = self._sv(item.get("hs_code", ""))
            desc  = self._sv(item.get("description", ""))
            net   = self._sv(item.get("net_weight", ""))
            tot   = self._sv(item.get("total_line", ""))
            extras = item.get("extra_lines", [])

            if hs:   lines_c2 += [f"HS. CODE. {hs}"]
            if desc: lines_c2 += simpleSplit(desc,  "Helvetica", 7.5, col2_w)
            if net:  lines_c2 += [f"NET WEIGHT {net}"]
            if tot:  lines_c2 += simpleSplit(tot,   "Helvetica", 7.5, col2_w)
            for e in extras:
                if e: lines_c2 += simpleSplit(e, "Helvetica", 7.5, col2_w)

            row_h   = max(ROW_MIN, len(lines_c2) * 12 + 16)
            pdf_bot = pdf_top + row_h

            # Draw column dividers and bottom rule
            for x in COL_X_PDF:
                self._line(can, x, _rl(pdf_top), x, _rl(pdf_bot))
            self._line(can, L_BORDER, _rl(pdf_bot), R_BORDER, _rl(pdf_bot))

            # Col 0 – Item No
            mid_col0 = (23 + 73) / 2
            self._centred(can, mid_col0, _rl(pdf_top + 14),
                          self._sv(item.get("item_no", "")))

            # Col 1 – Marks
            my_pdf = pdf_top + 8
            can.setFont("Helvetica", 7.5)
            for ln in self._sv(item.get("marks", "N/M")).split("\n"):
                if ln:
                    can.drawString(73 + 2, _rl(my_pdf), ln)
                    my_pdf += 12

            # Col 2 – HS / description / net weight / total / extras
            dy_pdf = pdf_top + 8
            can.setFont("Helvetica", 7.5)
            for ln in lines_c2:
                can.drawString(138 + 2, _rl(dy_pdf), ln)
                dy_pdf += 12

            # Col 3 – Origin Criterion
            mid_col3 = (314 + 385) / 2
            self._centred(can, mid_col3, _rl(pdf_top + 20),
                          self._sv(item.get("origin_criterion", "")))

            # Col 4 – Gross weight / FOB
            gwy_pdf = pdf_top + 8
            can.setFont("Helvetica", 7.5)
            gw  = self._sv(item.get("gross_weight", ""))
            fob = self._sv(item.get("fob_value", ""))
            if gw:
                can.drawString(385 + 4, _rl(gwy_pdf), gw);  gwy_pdf += 12
            if fob:
                can.drawString(385 + 4, _rl(gwy_pdf), fob)

            # Col 5 – Invoice No / Date
            iy_pdf = pdf_top + 8
            can.setFont("Helvetica", 7.5)
            inv_no   = self._sv(item.get("invoice_no", ""))
            inv_date = self._sv(item.get("invoice_date", ""))
            if inv_no:
                can.drawString(489 + 4, _rl(iy_pdf), inv_no);  iy_pdf += 12
            if inv_date:
                can.drawString(489 + 4, _rl(iy_pdf), inv_date)

            pdf_top = pdf_bot

        return pdf_top

    # ── Filler gap between last item row and Box 11 ───────────────────────────
    def _draw_filler(self, can, from_pdf_top: int, to_pdf_top: int):
        if from_pdf_top < to_pdf_top:
            self._line(can, L_BORDER, _rl(to_pdf_top), R_BORDER, _rl(to_pdf_top))
            for x in COL_X_PDF:
                self._line(can, x, _rl(from_pdf_top), x, _rl(to_pdf_top))

    # ── Boxes 11 + 12  (pdf_top 521 → 737) ───────────────────────────────────
    def _draw_boxes_11_12(self, can, pdf_top: int, data):
        hd   = data.get("header_data", {})
        decl = data.get("declaration", {})
        mid_l = (L_BORDER + MID_DIV) / 2
        bx    = L_BORDER + 6
        rx    = MID_DIV + 6

        # Vertical divider full height of the box
        self._line(can, MID_DIV, _rl(pdf_top), MID_DIV, _rl(737))

        # Sub-row horizontal rules (exact pdf_top values from reference)
        for r in [561, 589, 605, 633, 661, 677, 705, 721, 737]:
            self._line(can, L_BORDER, _rl(r), R_BORDER, _rl(r))

        # ── Box 11 ────────────────────────────────────────────────────────────
        self._bold(can, bx, _rl(pdf_top + 8), "11.Declaration by the exporter", size=7)

        can.setFont("Helvetica", 7.5)
        can.drawString(bx + 2, _rl(pdf_top + 20),
            "The undersigned hereby declares that the above details and statement are")
        can.drawString(bx + 2, _rl(pdf_top + 32),
            "correct; that all the goods were produced in")

        # Origin country + dotted line (row 561→589)
        origin = self._sv(decl.get("origin_country", "SINGAPORE"))
        self._centred(can, mid_l, _rl(561 + 10), origin, font="Helvetica-Bold", size=8)
        self._centred(can, mid_l, _rl(561 + 22), "...............................")

        # "and that they comply…" (row 605→633)
        can.setFont("Helvetica", 7.5)
        can.drawString(bx + 2, _rl(605 + 8),
            "and that they comply with the origin requirements specified for these goods in")
        can.drawString(bx + 2, _rl(605 + 20),
            "the ASEAN Trade in Goods Agreement for the goods exported to")

        # Destination country + dotted line (row 633→661)
        dest = self._sv(decl.get("destination_country", ""))
        self._centred(can, mid_l, _rl(633 + 10), dest, font="Helvetica-Bold", size=8)
        self._centred(can, mid_l, _rl(633 + 22), "...............................")

        # Issued-in + date + dotted line (row 677→705)
        issued   = self._sv(hd.get("issued_in", "SINGAPORE")).upper()
        sig_date = self._sv(decl.get("date", ""))
        self._centred(can, mid_l, _rl(677 + 10), f"{issued}  {sig_date}", size=8)
        self._centred(can, mid_l, _rl(677 + 22), "...............................")

        # Signature label (row 721→737)
        can.setFont("Helvetica", 7)
        can.drawString(bx + 2, _rl(721 + 10),
            "Place and date, signature of authorised signatory")

        # ── Box 12 ────────────────────────────────────────────────────────────
        self._bold(can, rx, _rl(pdf_top + 8), "12.Certification", size=7)

        can.setFont("Helvetica", 7.5)
        can.drawString(rx + 2, _rl(pdf_top + 20),
            "It is hereby certified, on the basis of control carried out, that the")
        can.drawString(rx + 2, _rl(pdf_top + 32),
            "declaration by the exporter is correct")



    
# ── Page dimensions ──────────────────────────────────────────────────────────
W, H = 595, 842   # A4 portrait in points

# ── Layout constants (from pdfplumber coordinate analysis) ───────────────────
LABEL_X  = 32     # left edge of label text
COLON_X  = 172    # colon column
VALUE_X  = 200    # left edge of value text
ROW_H    = 14     # vertical gap between rows (pt)
START_TOP = 94    # pdf_top of first data row
FONT_SIZE = 10    # body font size


def _rl(pdf_top: float) -> float:
    """Convert pdfplumber top-origin coordinate → ReportLab bottom-origin y."""
    return H - pdf_top


def _sv(val) -> str:
    """Safe string: return empty string for None / blank / placeholder."""
    if val is None:
        return ""
    if str(val).strip() in ("", "--Select--"):
        return ""
    return str(val).strip()


# ════════════════════════════════════════════════════════════════════════════
# DJANGO VIEW
# ════════════════════════════════════════════════════════════════════════════

class PrintCoo(APIView):

    def get(self, request, PermitId):
        try:
            # ── 1. Fetch header data ──────────────────────────────────────────
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT
                        MSGId,
                        PermitNumber,
                        CertificateNumber
                    FROM CommonHeaderTbl
                    WHERE PermitId = %s
                """, [PermitId])
                hrow = cursor.fetchone()

            if not hrow:
                return HttpResponse("Permit not found", status=404)

            (MSGId, PermitNumber, CertificateNumber) = hrow

            # ── 2. Assemble data dict (fill remaining fields later) ───────────
            data = {
                "submitted_by":         "",   # TODO: fill from your table
                "message_id":           _sv(MSGId),
                "cr_uei_no":            "",   # TODO: fill from your table
                "permit_no":            _sv(PermitNumber),
                "certificate_no":       _sv(CertificateNumber),
                "approval_date":        "",   # TODO: fill from your table
                "validity_start_date":  "",   # TODO: fill from your table
                "validity_end_date":    "",   # TODO: fill from your table
                "sc_approve_condition": "",   # TODO: fill from your table
            }

            # ── 5. Generate PDF ───────────────────────────────────────────────
            pdf_bytes = self._build_pdf(data)
            response  = HttpResponse(pdf_bytes, content_type="application/pdf")
            response["Content-Disposition"] = (
                f'attachment; filename="{_sv(CertificateNumber)}_COO_CERTIFICATE.pdf"'
            )
            return response

        except Exception as e:
            import traceback
            traceback.print_exc()
            return HttpResponse(
                f"Error generating Certificate PDF: {str(e)}", status=500
            )

    # ══════════════════════════════════════════════════════════════════════════
    # PDF BUILDER
    # ══════════════════════════════════════════════════════════════════════════

    def _build_pdf(self, data: dict) -> bytes:
        """
        Build and return the certificate PDF bytes.

        data dict keys
        ──────────────
        submitted_by         – e.g. "KTTSG01"
        message_id           – e.g. "202107020013"
        cr_uei_no            – e.g. "201834618Z"
        permit_no            – e.g. "OD1G034141D"
        certificate_no       – e.g. "20219006227"
        approval_date        – e.g. "03/07/2021"
        validity_start_date  – e.g. "03/07/2021"
        validity_end_date    – e.g. "15/07/2021"
        sc_approve_condition – e.g. "APPROVED BY SINGAPORE CUSTOMS."
        """
        packet = io.BytesIO()
        can    = canvas.Canvas(packet, pagesize=(W, H))

        self._draw_page_number(can)
        self._draw_heading(can)
        self._draw_fields(can, data)

        can.save()
        packet.seek(0)
        return packet.read()

    # ── Section: PG : 1 OF 1 ─────────────────────────────────────────────────
    def _draw_page_number(self, can):
        """Top-right page counter, Times-Roman 10, matching pdf_top ≈ 17."""
        can.setFont("Times-Roman", 10)
        can.drawString(490, _rl(17 + 10), "PG : 1 OF 1")

    # ── Section: DETAILS FOR COO heading ─────────────────────────────────────
    def _draw_heading(self, can):
        """Bold heading at pdf_top ≈ 52, font size 12."""
        can.setFont("Helvetica-Bold", 12)
        can.drawString(LABEL_X, _rl(52 + 12), "DETAILS FOR COO")

    # ── Section: field rows + end-of-permit line ──────────────────────────────
    def _draw_fields(self, can, data: dict):
        """
        Render label / colon / value rows then the closing line.

        Field order matches the reference template exactly.
        The intentional typo "CERFITICATE NO" is preserved to match
        the original document format.
        """
        fields = [
            # (label,                    data_key)
            ("SUBMITTED BY",         "submitted_by"),
            ("MESSAGE ID",           "message_id"),
            ("CR UEI NO",            "cr_uei_no"),
            ("PERMIT NO",            "permit_no"),
            ("CERFITICATE NO",       "certificate_no"),   # original spelling kept
            ("APPROVAL DATE",        "approval_date"),
            ("VALIDITY START DATE",  "validity_start_date"),
            ("VALIDITY END DATE",    "validity_end_date"),
            ("SC APPROVE CONDITION", "sc_approve_condition"),
        ]

        can.setFont("Helvetica", FONT_SIZE)

        for i, (label, key) in enumerate(fields):
            y = _rl(START_TOP + FONT_SIZE + i * ROW_H)
            can.drawString(LABEL_X, y, label)
            can.drawString(COLON_X, y, ":")
            can.drawString(VALUE_X, y, _sv(data.get(key, "")))

        # "END OF CARGO CLEARANCE PERMIT." — one row below last field
        end_y = _rl(START_TOP + FONT_SIZE + len(fields) * ROW_H)
        can.drawString(
            VALUE_X, end_y,
            data.get("end_of_cargo_msg", "END OF CARGO CLEARANCE PERMIT.")
        )


class TransmitOutpayment(APIView):
    def post(self, request):
        try:
            permit_ids       = request.data.get("permitIds", [])
            declaration_type = request.data.get("declarationType", "")
            username         = request.data.get("user") or request.session.get("Username")
            touch_time       = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # ── Validation ─────────────────────────────────────────────────
            if not permit_ids:
                return Response({"error": "No permits selected"}, status=400)
            if not declaration_type:
                return Response({"error": "Declaration type is required"}, status=400)
            if not username:
                return Response({"error": "User is required"}, status=400)

            now      = datetime.now()
            ref_date = now.strftime("%Y%m%d")
            job_date = now.strftime("%y%m%d")

            copied_permits = []
            warnings = []  # collect issues to return in response for debugging

            with transaction.atomic():
                cursor = connection.cursor()

                # ── Get user account details ───────────────────────────────
                cursor.execute(
                    "SELECT AccountId, MailBoxId FROM ManageUser WHERE UserName = %s",
                    [username]
                )
                row = cursor.fetchone()
                if not row:
                    return Response({"error": f"User '{username}' not found"}, status=404)
                account_id, mailbox_id = row

                # GLOBAL starting count for JobId/MsgId
                cursor.execute(
                    """
                    SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                    FROM CommonHeaderTbl
                    WHERE JobId LIKE %s
                    """,
                    [f"K{job_date}%"]
                )
                job_count = cursor.fetchone()[0]

                # Per-user starting count for PermitId/RefId
                cursor.execute(
                    """
                    SELECT ISNULL(COUNT(*), 0) + 1 AS NextSeq
                    FROM CommonHeaderTbl
                    WHERE PermitId LIKE %s
                    """,
                    [f"{username}{ref_date}%"]
                )
                ref_count = cursor.fetchone()[0]

                for permit_id in permit_ids:

                    # ── Check permit exists ────────────────────────────────
                    cursor.execute(
                        "SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s",
                        [permit_id]
                    )
                    if not cursor.fetchone():
                        warnings.append(f"{permit_id}: source permit not found, skipped")
                        continue

                    # ── Diagnostic: check source has invoice/item rows ──────
                    cursor.execute(
                        "SELECT COUNT(*) FROM CommonInvoiceDtl WHERE PermitId = %s",
                        [permit_id]
                    )
                    src_invoice_count = cursor.fetchone()[0]

                    cursor.execute(
                        "SELECT COUNT(*) FROM CommonItemDtl WHERE PermitId = %s",
                        [permit_id]
                    )
                    src_item_count = cursor.fetchone()[0]

                    print(f"[TransmitOutpayment] Source {permit_id}: "
                          f"{src_invoice_count} invoice row(s), {src_item_count} item row(s)")

                    if src_invoice_count == 0:
                        warnings.append(f"{permit_id}: source has 0 invoice rows in CommonInvoiceDtl")
                    if src_item_count == 0:
                        warnings.append(f"{permit_id}: source has 0 item rows in CommonItemDtl")

                    ref_id        = f"{ref_count:03d}"
                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    cursor.execute(
                        """
                        INSERT INTO CommonHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType,
                            PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        )
                        SELECT
                            %s, %s, %s, %s, %s,
                            'OUTDEC', %s,
                            PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            NULL, NULL,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            'DRF', %s, %s,
                            NULL, 'NEW',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb,
                            NULL,           -- DeclarningFor: NULL to avoid truncation error
                            MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                        """,
                        [
                            ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                            declaration_type,
                            username, touch_time,
                            permit_id
                        ]
                    )

                    # ── Copy all child tables ──────────────────────────────
                    child_tables = {
                        "CommonInvoiceDtl": [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator",
                            "SupplierImporterRelationship", "SupplierCode", "ImportPartyCode",
                            "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
                        ],
                        "CommonItemDtl": [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "EndUserDescription", "Brand", "Model",
                            "InHAWBOBL", "OutHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode",
                            "GSTRate", "GSTUOM", "GSTAmount",
                            "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "LSPValue", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "CerItemQty", "CerItemUOM", "CIFValOfCer", "ManufactureCostDate",
                            "TexCat", "TexQuotaQty", "TexQuotaUOM", "CerInvNo", "CerInvDate",
                            "OriginOfCer", "HSCodeCer", "PerContent", "CertificateDescription",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                        ],
                        "CommonCASCDtl": [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
                            "RowNo", "CascCode1", "CascCode2", "CascCode3",
                            "MessageType", "TouchUser", "TouchTime", "EndUserDes", "CASCId"
                        ],
                        "CommonCPCDtl": [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ],
                        "CommonContainerDtl": [
                            "RowNo", "ContainerNo", "Size", "Weight",
                            "SealNo", "MessageType", "TouchUser", "TouchTime"
                        ],
                        "CommonFile": [
                            "Name", "ContentType", "Data", "DocumentType",
                            "TouchUser", "TouchTime", "filePath", "Size", "Type"
                        ],
                        "CommonPMT": [
                            "ConditionCode", "ConditionDesc", "PermitNumber",
                        ],
                    }

                    for table, cols in child_tables.items():
                        col_list    = ", ".join(["PermitId"] + cols)
                        select_cols = ", ".join(cols)
                        try:
                            cursor.execute(
                                f"""
                                INSERT INTO {table} ({col_list})
                                SELECT %s, {select_cols}
                                FROM {table}
                                WHERE PermitId = %s
                                """,
                                [new_permit_id, permit_id]
                            )
                            rows_inserted = cursor.rowcount
                            print(f"[TransmitOutpayment] {table}: {rows_inserted} row(s) "
                                  f"copied from {permit_id} -> {new_permit_id}")

                            if rows_inserted == 0 and table in ("CommonInvoiceDtl", "CommonItemDtl"):
                                warnings.append(
                                    f"{permit_id}: 0 rows copied into {table} "
                                    f"(source had {src_invoice_count if table=='CommonInvoiceDtl' else src_item_count} rows)"
                                )

                            if "MessageType" in cols:
                                cursor.execute(
                                    f"UPDATE {table} SET MessageType = 'OUTDEC' WHERE PermitId = %s",
                                    [new_permit_id]
                                )

                        except Exception as err:
                            print(f"[TransmitOutpayment] ERROR copying {table}: {err}")
                            warnings.append(f"{permit_id}: ERROR copying {table}: {str(err)}")
                            # Re-raise for the two critical tables so the real DB error surfaces
                            # and the transaction rolls back instead of silently continuing.
                            if table in ("CommonInvoiceDtl", "CommonItemDtl"):
                                raise

                    # ── Insert PermitCount ─────────────────────────────────
                    cursor.execute(
                        """
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'OUTDEC', %s, %s, %s, %s)
                        """,
                        [new_permit_id, account_id, msg_id, username, touch_time]
                    )

                    copied_permits.append(new_permit_id)

                    # Increment counters for next permit in this batch
                    job_count += 1
                    ref_count += 1

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) transmitted successfully as OutPayment",
                "copiedPermits": copied_permits,
                "declarationType": declaration_type,
                "messageType": "OUTDEC",
                "warnings": warnings,
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)



class CopyOutAmend(APIView):
    def post(self, request):
        copied_permits = []
        try:
            permits = request.data.get("permits", [])
            username = request.data.get("user")
            touch_time = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if not permits:
                return Response({"error": "No permits selected"}, status=400)
            if not username:
                return Response({"error": "User required"}, status=400)
            now = datetime.now()
            ref_date = now.strftime("%Y%m%d")
            job_date = now.strftime("%y%m%d")

            with transaction.atomic():
                cursor = connection.cursor()
                cursor.execute(
                    "SELECT AccountId, MailBoxId FROM ManageUser WHERE UserName = %s",
                    [username]
                )
                row = cursor.fetchone()
                if not row:
                    return Response({"error": f"User '{username}' not found"}, status=404)
                account_id, mailbox_id = row

                cursor.execute(
                    "SELECT ISNULL(COUNT(*), 0) + 1 AS Count FROM CommonHeaderTbl WHERE JobId LIKE %s",
                    [f"K{job_date}%"]
                )
                job_count = cursor.fetchone()[0]

                cursor.execute(
                    "SELECT ISNULL(COUNT(*), 0) + 1 AS NextSeq FROM CommonHeaderTbl WHERE PermitId LIKE %s",
                    [f"{username}{ref_date}%"]
                )
                ref_count = cursor.fetchone()[0]

                for permit_id in permits:
                    cursor.execute("SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s", [permit_id])
                    if not cursor.fetchone():
                        continue

                    ref_id        = f"{ref_count:03d}"
                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    # ── 1. Insert new CommonHeaderTbl row — force MessageType = OUTDEC, prmtStatus = AMD ──
                    cursor.execute("""
                        INSERT INTO CommonHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        )
                        SELECT
                            %s, %s, %s, %s, TradeNetMailboxID,
                            'OUTDEC', DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            'DRF', %s, %s,
                            PermitNumber, 'AMD',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror into OutHeaderTbl (same column set CopyOut already uses) ──
                    cursor.execute("""
                        INSERT INTO OutHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, EndUserCode, Manufacturer,
                            ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, EndUserCode, Manufacturer,
                            ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: Common(OLD) -> Common(NEW), then Common(NEW) -> Out* ──
                    child_tables = {
                        "CommonItemDtl": ("OutItemDtl", [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "EndUserDescription", "Brand", "Model",
                            "InHAWBOBL", "OutHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode",
                            "GSTRate", "GSTUOM", "GSTAmount",
                            "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "CerItemQty", "CerItemUOM", "CIFValOfCer", "ManufactureCostDate",
                            "TexCat", "TexQuotaQty", "TexQuotaUOM", "CerInvNo", "CerInvDate",
                            "OriginOfCer", "HSCodeCer", "PerContent", "CertificateDescription",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                        ]),
                        "CommonCASCDtl": ("OutCASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
                            "RowNo", "CascCode1", "CascCode2", "CascCode3",
                            "MessageType", "TouchUser", "TouchTime", "CASCId", "EndUserDes"
                        ]),
                        "CommonCPCDtl": ("OutCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("OutContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, (dst_table, cols) in child_tables.items():
                        col_list = ", ".join(cols)
                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── 4. Invoice (column-name rename ImportPartyCode -> ExportPartyCode) ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonInvoiceDtl (
                                PermitId, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime, ChkOtherInv
                            )
                            SELECT
                                %s, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime, ChkOtherInv
                            FROM CommonInvoiceDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonInvoiceDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO OutInvoiceDtl (
                                PermitId, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ExportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime
                            )
                            SELECT
                                %s, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime
                            FROM CommonInvoiceDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonInvoiceDtl -> OutInvoiceDtl: {err}")

                    # ── 5. File & PMT ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonFile: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO OutFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonFile -> OutFile: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO CommonPMT (PermitId, ConditionCode, ConditionDesc, PermitNumber)
                            SELECT %s, ConditionCode, ConditionDesc, PermitNumber
                            FROM CommonPMT WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying PMT: {err}")

                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'OUTDEC', %s, %s, %s, %s)
                    """, [new_permit_id, account_id, msg_id, username, touch_time])

                    copied_permits.append(new_permit_id)
                    job_count += 1
                    ref_count += 1

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) copied for amend successfully",
                "copiedPermits": copied_permits,
            })
        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": f"Database Error: {str(e)}"}, status=400)


class PostOutAmendTable(APIView):
    table = "CommonAmend"
    out_table = "OutAmend"
    allowed_columns = {
        "Permitno", "AmendmentCount", "UpdateIndicator",
        "ReplacementPermitno", "DescriptionOfReason",
        "PermitExtension", "ExtendImportPeriod",
        "DeclarationIndigator", "AmendType",
        "TouchUser", "TouchTime", "MSGId",
    }

    OUT_TABLE_COLUMN_MAP = {
        "Permitno": "Permitno",
        "AmendmentCount": "AmendmentCount",
        "UpdateIndicator": "UpdateIndicator",
        "ReplacementPermitno": "ReplacementPermitno",
        "DescriptionOfReason": "DescriptionOfReason",
        "PermitExtension": "PermitExtension",
        "ExtendImportPeriod": "ExtendImportPeriod",
        "DeclarationIndigator": "DeclarationIndigator",
        "TouchUser": "TouchUser",
        "TouchTime": "TouchTme",  
        "MSGId": "MSGId",
        "AmendType": "AmendType",
    }

    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        item.pop("Id", None)
        msg_id = item.get("MSGId", "")

        if not msg_id:
            return Response({"error": "MSGId is required"}, status=400)

        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns provided"}, status=400)

        try:
            existing = SqlDb.execute_query(
                f"SELECT Id FROM {self.table} WHERE MSGId = %s", [msg_id]
            )
            if existing:
                update_cols = [c for c in columns if c != "MSGId"]
                set_clause = ", ".join([f"{c} = %s" for c in update_cols])
                values = [item[c] for c in update_cols] + [msg_id]
                SqlDb.execute_query(
                    f"UPDATE {self.table} SET {set_clause} WHERE MSGId = %s", values
                )
                action = "updated"
            else:
                col_str = ", ".join(columns)
                ph_str = ", ".join(["%s"] * len(columns))
                values = [item[c] for c in columns]
                SqlDb.execute_query(
                    f"INSERT INTO {self.table} ({col_str}) VALUES ({ph_str})", values
                )
                action = "inserted"

            out_item = {}
            for src_col, dst_col in self.OUT_TABLE_COLUMN_MAP.items():
                if src_col in item:
                    out_item[dst_col] = item.get(src_col)

            if out_item:
                existing_out = SqlDb.execute_query(
                    f"SELECT Id FROM {self.out_table} WHERE MSGId = %s", [msg_id]
                )
                if existing_out:
                    out_update_cols = [c for c in out_item.keys() if c != "MSGId"]
                    if out_update_cols:
                        out_set_clause = ", ".join([f"{c} = %s" for c in out_update_cols])
                        out_values = [out_item[c] for c in out_update_cols] + [msg_id]
                        SqlDb.execute_query(
                            f"UPDATE {self.out_table} SET {out_set_clause} WHERE MSGId = %s",
                            out_values
                        )
                else:
                    out_col_str = ", ".join(out_item.keys())
                    out_ph_str = ", ".join(["%s"] * len(out_item))
                    out_values = list(out_item.values())
                    SqlDb.execute_query(
                        f"INSERT INTO {self.out_table} ({out_col_str}) VALUES ({out_ph_str})",
                        out_values
                    )

            SqlDb.commit()

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response(
            {"Result": f"Amend record {action} successfully", "MSGId": msg_id},
            status=201
        )


# Post Cancel

class PostOutCancelTable(APIView):
    table = "CommonCancel"
    out_table = "OutCancel"
    allowed_columns = {
        "Permitno", "UpdateIndicator", "ReplacementPermitno", "ReasonForCancel",
        "DescriptionOfReason", "DeclarationIndigator",
        "TouchUser", "TouchTime", "MSGId", "CancelType",
    }

    # CommonCancel column -> OutCancel column
    # (OutCancel has typo'd column names: ResonForCancel, TouchTme — same as InNonCancel)
    OUT_TABLE_COLUMN_MAP = {
        "Permitno": "Permitno",
        "UpdateIndicator": "UpdateIndicator",
        "ReplacementPermitno": "ReplacementPermitno",
        "ReasonForCancel": "ResonForCancel",
        "DescriptionOfReason": "DescriptionOfReason",
        "DeclarationIndigator": "DeclarationIndigator",
        "TouchUser": "TouchUser",
        "TouchTime": "TouchTme",
        "MSGId": "MSGId",
        "CancelType": "CancelType",
    }

    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        item.pop("Id", None)
        msg_id = item.get("MSGId", "")
        permit_id = item.get("Permitno", "")

        if not msg_id:
            return Response({"error": "MSGId is required"}, status=400)

        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns provided"}, status=400)

        try:
            # ── 1. CommonCancel upsert ──
            existing = SqlDb.execute_query(
                f"SELECT Id FROM {self.table} WHERE MSGId = %s", [msg_id]
            )
            if existing:
                update_cols = [c for c in columns if c != "MSGId"]
                set_clause = ", ".join([f"{c} = %s" for c in update_cols])
                values = [item[c] for c in update_cols] + [msg_id]
                SqlDb.execute_query(
                    f"UPDATE {self.table} SET {set_clause} WHERE MSGId = %s", values
                )
                action = "updated"
            else:
                col_str = ", ".join(columns)
                ph_str = ", ".join(["%s"] * len(columns))
                values = [item[c] for c in columns]
                SqlDb.execute_query(
                    f"INSERT INTO {self.table} ({col_str}) VALUES ({ph_str})", values
                )
                action = "inserted"

            # ── 2. Mirror into OutCancel (remapped column names) ──
            out_item = {}
            for src_col, dst_col in self.OUT_TABLE_COLUMN_MAP.items():
                if src_col in item:
                    out_item[dst_col] = item.get(src_col)

            if out_item:
                existing_out = SqlDb.execute_query(
                    f"SELECT Id FROM {self.out_table} WHERE MSGId = %s", [msg_id]
                )
                if existing_out:
                    out_update_cols = [c for c in out_item.keys() if c != "MSGId"]
                    if out_update_cols:
                        out_set_clause = ", ".join([f"{c} = %s" for c in out_update_cols])
                        out_values = [out_item[c] for c in out_update_cols] + [msg_id]
                        SqlDb.execute_query(
                            f"UPDATE {self.out_table} SET {out_set_clause} WHERE MSGId = %s",
                            out_values
                        )
                else:
                    out_col_str = ", ".join(out_item.keys())
                    out_ph_str = ", ".join(["%s"] * len(out_item))
                    out_values = list(out_item.values())
                    SqlDb.execute_query(
                        f"INSERT INTO {self.out_table} ({out_col_str}) VALUES ({out_ph_str})",
                        out_values
                    )

            SqlDb.commit()

            # ── 3. Update status on BOTH header tables ──
            if permit_id:
                SqlDb.execute_query(
                    "UPDATE CommonHeaderTbl SET prmtStatus = 'CNL' WHERE PermitId = %s",
                    [permit_id]
                )
                try:
                    SqlDb.execute_query(
                        "UPDATE OutHeaderTbl SET prmtStatus = 'CNL' WHERE PermitId = %s",
                        [permit_id]
                    )
                except Exception as err:
                    print(f"Warning updating OutHeaderTbl prmtStatus: {err}")
                SqlDb.commit()

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response({"Result": f"Cancel record {action} successfully", "MSGId": msg_id}, status=201)


# Copy Cancel
class CopyOutCancel(APIView):
    def post(self, request):
        copied_permits = []
        try:
            permits = request.data.get("permits", [])
            username = request.data.get("user")
            touch_time = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if not permits:
                return Response({"error": "No permits selected"}, status=400)
            if not username:
                return Response({"error": "User required"}, status=400)
            now = datetime.now()
            ref_date = now.strftime("%Y%m%d")
            job_date = now.strftime("%y%m%d")

            with transaction.atomic():
                cursor = connection.cursor()
                cursor.execute(
                    "SELECT AccountId, MailBoxId FROM ManageUser WHERE UserName = %s",
                    [username]
                )
                row = cursor.fetchone()
                if not row:
                    return Response({"error": f"User '{username}' not found"}, status=404)
                account_id, mailbox_id = row

                cursor.execute(
                    "SELECT ISNULL(COUNT(*), 0) + 1 AS Count FROM CommonHeaderTbl WHERE JobId LIKE %s",
                    [f"K{job_date}%"]
                )
                job_count = cursor.fetchone()[0]

                cursor.execute(
                    "SELECT ISNULL(COUNT(*), 0) + 1 AS NextSeq FROM CommonHeaderTbl WHERE PermitId LIKE %s",
                    [f"{username}{ref_date}%"]
                )
                ref_count = cursor.fetchone()[0]

                for permit_id in permits:
                    cursor.execute("SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s", [permit_id])
                    if not cursor.fetchone():
                        continue

                    ref_id        = f"{ref_count:03d}"
                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    # ── 1. Insert new CommonHeaderTbl row — copy everything as-is,
                    #      force MessageType to OUTDEC and prmtStatus to CNL ──
                    cursor.execute("""
                        INSERT INTO CommonHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        )
                        SELECT
                            %s, %s, %s, %s, TradeNetMailboxID,
                            'OUTDEC', DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, ClaimantPartyCode, EndUserCode, Manufacturer,
                            HBL, ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            'DRF', %s, %s,
                            PermitNumber, 'CNL',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror into OutHeaderTbl (same column set CopyOut/CopyOutAmend use) ──
                    cursor.execute("""
                        INSERT INTO OutHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, EndUserCode, Manufacturer,
                            ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            Inwardcarriercode, OutwardCarrierAgentCode,
                            FreightForwarderCode, ImporterCompanyCode, InwardCarrierAgentCode,
                            CONSIGNEECode, EndUserCode, Manufacturer,
                            ArrivalDate, ArrivalTime, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo,
                            ConveyanceRefNo, TransportId, FlightNO, AircraftRegNo,
                            MasterAirwayBill, ReleaseLocation, RecepitLocation,
                            StorageLocation, BlanketStartDate,
                            DepartureDate, DepartureTime, DischargePort,
                            FinalDestinationCountry, OutVoyageNumber, OutVesselName,
                            OutOceanBillofLadingNo, VesselType, VesselNetRegTon,
                            VesselNationality, TowingVesselID, TowingVesselName,
                            NextPort, LastPort, OutConveyanceRefNo, OutTransportId,
                            OutFlightNO, OutAircraftRegNo, OutMasterAirwayBill,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime,
                            PermitNumber, prmtStatus,
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: Common(OLD) -> Common(NEW), then Common(NEW) -> Out* ──
                    child_tables = {
                        "CommonItemDtl": ("OutItemDtl", [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "EndUserDescription", "Brand", "Model",
                            "InHAWBOBL", "OutHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode",
                            "GSTRate", "GSTUOM", "GSTAmount",
                            "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "CerItemQty", "CerItemUOM", "CIFValOfCer", "ManufactureCostDate",
                            "TexCat", "TexQuotaQty", "TexQuotaUOM", "CerInvNo", "CerInvDate",
                            "OriginOfCer", "HSCodeCer", "PerContent", "CertificateDescription",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                            # NOTE: LSPValue has no counterpart column in OutItemDtl — dropped.
                        ]),
                        "CommonCASCDtl": ("OutCASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
                            "RowNo", "CascCode1", "CascCode2", "CascCode3",
                            "MessageType", "TouchUser", "TouchTime", "CASCId", "EndUserDes"
                        ]),
                        "CommonCPCDtl": ("OutCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("OutContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, (dst_table, cols) in child_tables.items():
                        col_list = ", ".join(cols)
                        # Common -> Common (from OLD permit)
                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")

                        # Common(NEW) -> Out target — reads what we just wrote
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── 4. Invoice — ExportPartyCode rename, no ChkOtherInv on Out side ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonInvoiceDtl (
                                PermitId, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime, ChkOtherInv
                            )
                            SELECT
                                %s, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime, ChkOtherInv
                            FROM CommonInvoiceDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonInvoiceDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO OutInvoiceDtl (
                                PermitId, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ExportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime
                            )
                            SELECT
                                %s, SNo, InvoiceNo, InvoiceDate, TermType,
                                AdValoremIndicator, PreDutyRateIndicator, SupplierImporterRelationship,
                                SupplierCode, ImportPartyCode,
                                TICurrency, TIExRate, TIAmount, TISAmount,
                                OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                                FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                                ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                                CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                                MessageType, TouchUser, TouchTime
                            FROM CommonInvoiceDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonInvoiceDtl -> OutInvoiceDtl: {err}")

                    # ── 5. File — OutFile has no filePath ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonFile: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO OutFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonFile -> OutFile: {err}")

                    # ── 6. PMT — copy Common only ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonPMT (PermitId, ConditionCode, ConditionDesc, PermitNumber)
                            SELECT %s, ConditionCode, ConditionDesc, PermitNumber
                            FROM CommonPMT WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying PMT: {err}")

                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'OUTDEC', %s, %s, %s, %s)
                    """, [new_permit_id, account_id, msg_id, username, touch_time])

                    copied_permits.append(new_permit_id)
                    job_count += 1
                    ref_count += 1

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) copied for cancel successfully",
                "copiedPermits": copied_permits,
            })
        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": f"Database Error: {str(e)}"}, status=400)