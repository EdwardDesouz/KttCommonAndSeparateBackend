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


class TransList(APIView):
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
                    i.Name + ' ' + i.Name1 AS IMPORTER,
                    t1.HBL AS HAWB,
                    CASE  
                        WHEN t1.InwardTransportMode = '4 : Air' THEN t1.MasterAirwayBill  
                        WHEN t1.InwardTransportMode = '1 : Sea' THEN t1.OceanBillofLadingNo  
                        ELSE '' 
                    END AS MAWBOBL,
                    t1.LoadingPortCode AS POL,
                    t1.MessageType AS MSGTYPE,
                    t1.InwardTransportMode AS TPT,
                    t1.PreviousPermit AS PREPMT,
                    t1.GrossReference AS XREF,
                    t1.InternalRemarks AS INTREM,
                    t1.Message As MSG,
                    t1.TotalGSTTaxAmt AS GSTAMT,
                    t1.Status,
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
                LEFT JOIN CommonImporter i ON t1.ImporterCompanyCode = i.Code
            """
            # ↑ ManageUser JOIN completely REMOVED

            if show_all:
                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'TNPDEC'
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId])

            else:
                nowdate = datetime.now() - timedelta(days=90)
                date_filter = nowdate.strftime("%Y/%m/%d")

                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'TNPDEC'
                    AND CONVERT(varchar, t1.TouchTime, 111) >= %s
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId, date_filter])

            return Response(result)

        except Exception as e:
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)


# New Permit
class TranshipmentNewPermit(APIView):
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

            # q_ref = """
            #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
            #     FROM CommonHeaderTbl 
            #     WHERE MSGId LIKE %s AND MessageType = 'INPDEC'
            # """
            # ref_rows = SqlDb.execute_query(q_ref, [f"%{refDate}%"])
            # ref_rows = SqlDb.execute_query(q_ref, [f"{refDate}%"])
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

            # q_job = """
            #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
            #     FROM PermitCount 
            #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'INPDEC'
            # """

            # job_rows = SqlDb.execute_query(
                # """
                # # SELECT ISNULL(MAX(CAST(SUBSTRING(JobId, 2, 6) + RIGHT(JobId, 5) AS BIGINT)), 0) + 1 AS Count
                # # FROM CommonHeaderTbl
                # """
                # 6-6-2026
            # job_rows = SqlDb.execute_query(
            #     """
            #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
            #     FROM PermitCount 
            #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'TNPDEC'
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
            # End
# 4. Company info
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





# class CopyTranshipment(APIView):
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
#                     # 6-6-26 start
#                     # cursor.execute("""
#                     #     SELECT ISNULL(COUNT(*), 0) + 1
#                     #     FROM PermitCount
#                     #     WHERE TouchTime LIKE %s AND AccountId = %s
#                     # """, [f"%{today_dash}%", account_id])
#                     # End
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
#                             CondColor, TransmitId, gstVerified,HandlingAgentCode
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
#                             CondColor, TransmitId, gstVerified,HandlingAgentCode
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
#                         VALUES (%s, 'TNPDEC', %s, %s, %s, %s)
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

# class CopyTranshipment(APIView):
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
#                             CondColor, TransmitId, gstVerified,HandlingAgentCode
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
#                             CondColor, TransmitId, gstVerified,HandlingAgentCode
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
#                         VALUES (%s, 'TNPDEC', %s, %s, %s, %s)
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

class CopyTranshipment(APIView):
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
                            CondColor, TransmitId, gstVerified, HandlingAgentCode
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
                            Defrentprinting, Cnb, NULL, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified, HandlingAgentCode
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror the new row into TranshipmentHeader ──
                    # Note: ReleaseLocName <- ResLoaName, RemovalStartDate <- BlanketStartDate.
                    # Fields with no TranshipmentHeader counterpart (HBL, ArrivalTime,
                    # DepartureTime, seastore, gstVerified, Message, CustomerRemarks,
                    # ExporterCompanyCode, ConsigneeCode, etc.) are simply not mirrored.
                    cursor.execute("""
                        INSERT INTO TranshipmentHeader (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License, Recipient,
                            DeclarantCompanyCode, ImporterCompanyCode, HandlingAgentCode,
                            InwardCarrierAgentCode, OutwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, EndUserCode, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, RecepitLocation, StorageLocation,
                            RemovalStartDate, DepartureDate, DischargePort,
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
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            ReleaseLocName, RecepitLocName, Cnb, DeclarningFor,
                            INHAWB, outHAWB, MRDate, MRTime, CondColor, TransmitId
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License, Recipient,
                            DeclarantCompanyCode, ImporterCompanyCode, HandlingAgentCode,
                            InwardCarrierAgentCode, OutwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, EndUserCode, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, RecepitLocation, StorageLocation,
                            BlanketStartDate, DepartureDate, DischargePort,
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
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            ResLoaName, RecepitLocName, Cnb, DeclarningFor,
                            INHAWB, outHAWB, MRDate, MRTime, CondColor, TransmitId
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables that map 1:1 between Common* and Transhipment* ──
                    # (CPC and Container column names match exactly on both sides.)
                    child_tables = {
                        "CommonCPCDtl": ("TranshipmentCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("TranshipmentContainerDtl", [
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

                        # Common(NEW) -> Transhipment target
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── 4. Item — column names differ (VehicleType/EngineCapcity/
                    # EngineCapUOM/orignaldatereg on Common vs DrpVehicleType/
                    # Enginecapacity/Engineuom/Orginregdate on Transhipment). Also,
                    # TranshipmentItemDtl has no EndUserDescription, InvoiceNo,
                    # LSPValue, CerItemQty/UOM, CIFValOfCer, ManufactureCostDate,
                    # TexCat, TexQuotaQty/UOM, CerInvNo/Date, OriginOfCer, HSCodeCer,
                    # PerContent, CertificateDescription — those stay Common-only. ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonItemDtl (
                                PermitId, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, EndUserDescription, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer, InvoiceNo,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, LSPValue, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                CerItemQty, CerItemUOM, CIFValOfCer, ManufactureCostDate,
                                TexCat, TexQuotaQty, TexQuotaUOM, CerInvNo, CerInvDate,
                                OriginOfCer, HSCodeCer, PerContent, CertificateDescription,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            )
                            SELECT
                                %s, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, EndUserDescription, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer, InvoiceNo,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, LSPValue, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                CerItemQty, CerItemUOM, CIFValOfCer, ManufactureCostDate,
                                TexCat, TexQuotaQty, TexQuotaUOM, CerInvNo, CerInvDate,
                                OriginOfCer, HSCodeCer, PerContent, CertificateDescription,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            FROM CommonItemDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonItemDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO TranshipmentItemDtl (
                                PermitId, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                TouchUser, TouchTime, DrpVehicleType, OptionalChrgeUOM,
                                Enginecapacity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, Engineuom, Orginregdate
                            )
                            SELECT
                                %s, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            FROM CommonItemDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonItemDtl -> TranshipmentItemDtl: {err}")

                    # ── 5. CASC — column name differs (EndUserDes vs Enduserdesc) ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonCASCDtl (
                                PermitId, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, EndUserDes, CASCId
                            )
                            SELECT
                                %s, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, EndUserDes, CASCId
                            FROM CommonCASCDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonCASCDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO TCASCDtl (
                                PermitId, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, CASCId, Enduserdesc
                            )
                            SELECT
                                %s, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, CASCId, EndUserDes
                            FROM CommonCASCDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonCASCDtl -> TCASCDtl: {err}")

                    # ── 6. File — transhipfile has no filePath, uses Touchuser
                    # (lowercase u) and no PaymentId/PermitId-linked TranshipId ──
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
                            INSERT INTO transhipfile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, Touchuser, TouchTime, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonFile -> transhipfile: {err}")

                    # ── 7. PMT — copy Common only (no transhipment-specific PMT table given) ──
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
                        VALUES (%s, 'TNPDEC', %s, %s, %s, %s)
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


class PostTransAmendTable(APIView):
    table = "CommonAmend"
    out_table = "TransAmend"
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


class CopyTransAmend(APIView):
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

                    # ── 1. Insert new CommonHeaderTbl row — force MessageType = TNPDEC,
                    # keep PermitNumber (unlike plain copy which nulls it), prmtStatus = AMD ──
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
                            CondColor, TransmitId, gstVerified, HandlingAgentCode
                        )
                        SELECT
                            %s, %s, %s, %s, TradeNetMailboxID,
                            'TNPDEC', DeclarationType, PreviousPermit, CargoPackType,
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
                            CondColor, TransmitId, gstVerified, HandlingAgentCode
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror the new row into TranshipmentHeader ──
                    # (same column mapping as CopyTranshipment: ReleaseLocName <- ResLoaName,
                    # RemovalStartDate <- BlanketStartDate)
                    cursor.execute("""
                        INSERT INTO TranshipmentHeader (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License, Recipient,
                            DeclarantCompanyCode, ImporterCompanyCode, HandlingAgentCode,
                            InwardCarrierAgentCode, OutwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, EndUserCode, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, RecepitLocation, StorageLocation,
                            RemovalStartDate, DepartureDate, DischargePort,
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
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            ReleaseLocName, RecepitLocName, Cnb, DeclarningFor,
                            INHAWB, outHAWB, MRDate, MRTime, CondColor, TransmitId
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License, Recipient,
                            DeclarantCompanyCode, ImporterCompanyCode, HandlingAgentCode,
                            InwardCarrierAgentCode, OutwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, EndUserCode, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, RecepitLocation, StorageLocation,
                            BlanketStartDate, DepartureDate, DischargePort,
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
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            ResLoaName, RecepitLocName, Cnb, DeclarningFor,
                            INHAWB, outHAWB, MRDate, MRTime, CondColor, TransmitId
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables — CPC & Container map 1:1 between Common* and Transhipment* ──
                    child_tables = {
                        "CommonCPCDtl": ("TranshipmentCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("TranshipmentContainerDtl", [
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

                    # ── 4. Item — column-name rename between Common and Transhipment ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonItemDtl (
                                PermitId, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, EndUserDescription, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer, InvoiceNo,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, LSPValue, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                CerItemQty, CerItemUOM, CIFValOfCer, ManufactureCostDate,
                                TexCat, TexQuotaQty, TexQuotaUOM, CerInvNo, CerInvDate,
                                OriginOfCer, HSCodeCer, PerContent, CertificateDescription,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            )
                            SELECT
                                %s, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, EndUserDescription, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer, InvoiceNo,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, LSPValue, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                CerItemQty, CerItemUOM, CIFValOfCer, ManufactureCostDate,
                                TexCat, TexQuotaQty, TexQuotaUOM, CerInvNo, CerInvDate,
                                OriginOfCer, HSCodeCer, PerContent, CertificateDescription,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            FROM CommonItemDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonItemDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO TranshipmentItemDtl (
                                PermitId, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                TouchUser, TouchTime, DrpVehicleType, OptionalChrgeUOM,
                                Enginecapacity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, Engineuom, Orginregdate
                            )
                            SELECT
                                %s, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            FROM CommonItemDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonItemDtl -> TranshipmentItemDtl: {err}")

                    # ── 5. CASC — column name differs (EndUserDes vs Enduserdesc) ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonCASCDtl (
                                PermitId, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, EndUserDes, CASCId
                            )
                            SELECT
                                %s, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, EndUserDes, CASCId
                            FROM CommonCASCDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonCASCDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO TCASCDtl (
                                PermitId, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, CASCId, Enduserdesc
                            )
                            SELECT
                                %s, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, CASCId, EndUserDes
                            FROM CommonCASCDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonCASCDtl -> TCASCDtl: {err}")

                    # ── 6. File & PMT ──
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
                            INSERT INTO transhipfile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, Touchuser, TouchTime, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonFile -> transhipfile: {err}")

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
                        VALUES (%s, 'TNPDEC', %s, %s, %s, %s)
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

# POST TRANS CANCEL



class PostCancelTranshipment(APIView):
    table = "CommonCancel"
    out_table = "TransCancel"
    allowed_columns = {
        "Permitno", "UpdateIndicator", "ReplacementPermitno", "ReasonForCancel",
        "DescriptionOfReason", "DeclarationIndigator",
        "TouchUser", "TouchTime", "MSGId", "CancelType",
    }

    # CommonCancel column -> TransCancel column
    # (TransCancel keeps the typo'd ResonForCancel / TouchTme, per schema dump)
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
            # ── 1. CommonCancel upsert ──────────────────────────────────
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

            # ── 2. Mirror into TransCancel (remapped column names) ─────
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

            # ── 3. Update status on BOTH header tables ──────────────────
            if permit_id:
                SqlDb.execute_query(
                    "UPDATE CommonHeaderTbl SET prmtStatus = 'CNL' WHERE PermitId = %s",
                    [permit_id]
                )
                try:
                    SqlDb.execute_query(
                        "UPDATE TranshipmentHeader SET prmtStatus = 'CNL' WHERE PermitId = %s",
                        [permit_id]
                    )
                except Exception as err:
                    print(f"Warning updating TranshipmentHeader prmtStatus: {err}")
                SqlDb.commit()

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response(
            {"Result": f"Cancel record {action} successfully", "MSGId": msg_id},
            status=201
        )




class CopyCancelTranshipment(APIView):
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

                    # ── 1. Insert new CommonHeaderTbl row — keep PermitNumber
                    # (unlike plain CopyTranshipment which nulls it), force
                    # MessageType = TNPDEC, prmtStatus = CNL ──
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
                            CondColor, TransmitId, gstVerified, HandlingAgentCode
                        )
                        SELECT
                            %s, %s, %s, %s, TradeNetMailboxID,
                            'TNPDEC', DeclarationType, PreviousPermit, CargoPackType,
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
                            CondColor, TransmitId, gstVerified, HandlingAgentCode
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror the new row into TranshipmentHeader ──
                    # (same column mapping as CopyTranshipment/CopyTransAmend:
                    # ReleaseLocName <- ResLoaName, RemovalStartDate <- BlanketStartDate)
                    cursor.execute("""
                        INSERT INTO TranshipmentHeader (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License, Recipient,
                            DeclarantCompanyCode, ImporterCompanyCode, HandlingAgentCode,
                            InwardCarrierAgentCode, OutwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, EndUserCode, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, RecepitLocation, StorageLocation,
                            RemovalStartDate, DepartureDate, DischargePort,
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
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            ReleaseLocName, RecepitLocName, Cnb, DeclarningFor,
                            INHAWB, outHAWB, MRDate, MRTime, CondColor, TransmitId
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License, Recipient,
                            DeclarantCompanyCode, ImporterCompanyCode, HandlingAgentCode,
                            InwardCarrierAgentCode, OutwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, EndUserCode, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, RecepitLocation, StorageLocation,
                            BlanketStartDate, DepartureDate, DischargePort,
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
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            ResLoaName, RecepitLocName, Cnb, DeclarningFor,
                            INHAWB, outHAWB, MRDate, MRTime, CondColor, TransmitId
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables — CPC & Container map 1:1 between
                    # Common* and Transhipment* ──
                    child_tables = {
                        "CommonCPCDtl": ("TranshipmentCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("TranshipmentContainerDtl", [
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

                    # ── 4. Item — column-name rename between Common and Transhipment ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonItemDtl (
                                PermitId, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, EndUserDescription, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer, InvoiceNo,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, LSPValue, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                CerItemQty, CerItemUOM, CIFValOfCer, ManufactureCostDate,
                                TexCat, TexQuotaQty, TexQuotaUOM, CerInvNo, CerInvDate,
                                OriginOfCer, HSCodeCer, PerContent, CertificateDescription,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            )
                            SELECT
                                %s, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, EndUserDescription, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer, InvoiceNo,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, LSPValue, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                CerItemQty, CerItemUOM, CIFValOfCer, ManufactureCostDate,
                                TexCat, TexQuotaQty, TexQuotaUOM, CerInvNo, CerInvDate,
                                OriginOfCer, HSCodeCer, PerContent, CertificateDescription,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            FROM CommonItemDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonItemDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO TranshipmentItemDtl (
                                PermitId, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                TouchUser, TouchTime, DrpVehicleType, OptionalChrgeUOM,
                                Enginecapacity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, Engineuom, Orginregdate
                            )
                            SELECT
                                %s, ItemNo, MessageType, HSCode, Description, DGIndicator,
                                Contry, Brand, Model, InHAWBOBL, OutHAWBOBL,
                                DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                                InvoiceQuantity, HSQty, HSUOM, AlcoholPer,
                                ChkUnitPrice, UnitPrice, UnitPriceCurrency, ExchangeRate,
                                SumExchangeRate, TotalLineAmount, InvoiceCharges, CIFFOB,
                                OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                                PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                                ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                                CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                                OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                                CurrentLot, PreviousLot, Making,
                                ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                                TouchUser, TouchTime, VehicleType, OptionalChrgeUOM,
                                EngineCapcity, Optioncahrge, OptionalSumtotal,
                                OptionalSumExchage, EngineCapUOM, orignaldatereg
                            FROM CommonItemDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonItemDtl -> TranshipmentItemDtl: {err}")

                    # ── 5. CASC — column name differs (EndUserDes vs Enduserdesc) ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonCASCDtl (
                                PermitId, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, EndUserDes, CASCId
                            )
                            SELECT
                                %s, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, EndUserDes, CASCId
                            FROM CommonCASCDtl WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                    except Exception as err:
                        print(f"Warning copying CommonCASCDtl: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO TCASCDtl (
                                PermitId, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, CASCId, Enduserdesc
                            )
                            SELECT
                                %s, ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                MessageType, TouchUser, TouchTime, CASCId, EndUserDes
                            FROM CommonCASCDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonCASCDtl -> TCASCDtl: {err}")

                    # ── 6. File & PMT ──
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
                            INSERT INTO transhipfile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, Touchuser, TouchTime, Size, Type)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonFile -> transhipfile: {err}")

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
                        VALUES (%s, 'TNPDEC', %s, %s, %s, %s)
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