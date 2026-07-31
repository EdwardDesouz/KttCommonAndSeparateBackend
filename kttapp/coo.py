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



class CooList(APIView):
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
                    CONVERT(varchar, t1.DepartureDate, 105) AS ETD,
                    t1.PermitNumber AS PERMITNO,
                    i.Name + ' ' + i.Name1 AS EXPORTER,
                    t1.HBL AS HAWB,
                    CASE  
                        WHEN t1.InwardTransportMode = '4 : Air' THEN t1.MasterAirwayBill  
                        WHEN t1.InwardTransportMode = '1 : Sea' THEN t1.OceanBillofLadingNo  
                        ELSE '' 
                    END AS MAWBOBL,
                    t1.DischargePort AS POD,
                    CASE 
                    WHEN CHARINDEX(':', t1.COType) > 0 
                    THEN RTRIM(SUBSTRING(t1.COType, 1, CHARINDEX(':', t1.COType) - 1))
                    ELSE t1.COType 
                    END AS CoType,
                    CASE 
                    WHEN CHARINDEX(':', t1.CerDetailtype1) > 0 
                    THEN RTRIM(SUBSTRING(t1.CerDetailtype1, 1, CHARINDEX(':', t1.CerDetailtype1) - 1))
                    ELSE t1.CerDetailtype1 
                    END AS CerDetailtype1,
                    t1.MessageType AS MSGTYPE,
                    t1.OutwardTransportMode AS TPT,
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
                LEFT JOIN CommonExporter i ON t1.ExporterCompanyCode = i.Code
            """
            # ↑ ManageUser JOIN completely REMOVED

            if show_all:
                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'COODEC'
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId])

            else:
                nowdate = datetime.now() - timedelta(days=90)
                date_filter = nowdate.strftime("%Y/%m/%d")

                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'COODEC'
                    AND CONVERT(varchar, t1.TouchTime, 111) >= %s
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId, date_filter])

            return Response(result)

        except Exception as e:
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)


# # New Permit
# class CooNewPermit(APIView):
#     print('hello')
#     def get(self, request):
#         try:
#             Username = request.query_params.get("user")
#             if not Username:
#                 Username = request.session.get("Username")
#             if not Username:
#                 return Response({"error": "Session expired or User not provided"}, status=401)
#             refDate = datetime.now().strftime("%Y%m%d")
#             yy_mmdd = datetime.now().strftime("%Y-%m-%d")
#             currentDate = datetime.now().strftime("%d/%m/%Y")  

#             q_account = "SELECT AccountId FROM ManageUser WHERE UserName = %s"

#             account_rows = SqlDb.execute_query(q_account, [Username])
#             if not account_rows:
#                 return Response({"error": "User not found"}, status=404)
#             AccountId = account_rows[0]['AccountId']


# # 6-6-2026 Start
#             # ref_rows = SqlDb.execute_query(
#             #     """
#             #     SELECT ISNULL(COUNT(*), 0) + 1 AS Count
#             #     FROM CommonHeaderTbl
#             #     WHERE PermitId LIKE %s
#             #     """,
#             #     [f"{Username}{refDate}%"]
#             # )
           
           
           
#             # RefId = "%03d" % (ref_rows[0]['Count'] if ref_rows else 1)

           
           
#             # RefId = "%03d" % (ref_rows[0]['Count'] if ref_rows else 1)
# # End

#             count_rows = SqlDb.execute_query(
#                 """
#                 SELECT ISNULL(COUNT(*), 0) + 1 AS Count
#                 FROM CommonHeaderTbl
#                 WHERE PermitId LIKE %s
#                 """,
#                 [f"{Username}{refDate}%"]
#             )
#             count = count_rows[0]['Count'] if count_rows else 1

#             # 3. All 3 IDs from same count
#             RefId    = f"{count:03d}"
#             PermitId = f"{Username}{refDate}{RefId}"
#             JobId    = f"K{yy_mmdd}{count:05d}"
#             MsgId    = f"{refDate}{count:04d}"

#             print("PermitId:", PermitId, "| JobId:", JobId, "| MsgId:", MsgId, "| count:", count)


#             # job_rows = SqlDb.execute_query(
#             #     """
#             #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
#             #     FROM PermitCount 
#             #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'COODEC'
#             #     """,
#             #     [f"{jobDate}%", AccountId] 
#             # )
#             # JobIdCount = job_rows[0]['Count'] if job_rows else 1
#             # JobId = f"K{datetime.now().strftime('%y%m%d')}{JobIdCount:05d}"
            
#             # mailbox_rows = SqlDb.execute_query(
#             #     "SELECT MailBoxId FROM ManageUser WHERE UserName = %s",
#             #     [Username]
#             # )
#             # MailBoxId = mailbox_rows[0]['MailBoxId'] if mailbox_rows else ""
#             # msg_rows = SqlDb.execute_query(
#             #     """
#             #     SELECT ISNULL(MAX(CAST(RIGHT(MsgId, 4) AS INT)), 0) + 1 AS Count
#             #     FROM PermitCount
#             #     WHERE AccountId = %s
#             #     """,
#             #     [AccountId]
#             # )
#             # MsgCount = msg_rows[0]['Count'] if msg_rows else 1

#             # MsgId = f"{datetime.now().strftime('%Y%m%d')}{MsgCount:04d}"

#             # PermitId = f"{Username}{refDate}{RefId}"

#             # print("PermitId:", PermitId)
#             # print("JobId:", JobId)
#             # print("MsgId:", MsgId)
#             # print("RefId:", RefId)
#             # print('AccountId:', AccountId)

#             query_join = """
#                 SELECT TOP 1 
#                     manageuser.LoginStatus, manageuser.DateLastUpdated, manageuser.MailBoxId, 
#                     manageuser.SeqPool, SequencePool.StartSequence, DeclarantCompany.TradeNetMailboxID, 
#                     DeclarantCompany.DeclarantName, DeclarantCompany.DeclarantCode, 
#                     DeclarantCompany.DeclarantTel, DeclarantCompany.CRUEI, DeclarantCompany.Code, 
#                     DeclarantCompany.name, DeclarantCompany.name1 
#                 FROM manageuser 
#                 INNER JOIN SequencePool ON manageuser.SeqPool = SequencePool.Description 
#                 INNER JOIN DeclarantCompany ON DeclarantCompany.TradeNetMailboxID = ManageUser.MailBoxId 
#                 WHERE ManageUser.UserName = %s
#             """
#             head_rows = SqlDb.execute_query(query_join, [Username])

#             if not head_rows:
#                 return Response({"error": "Company profile data not found"}, status=404)

#             head = head_rows[0]

#             return Response({
#                 "UserName": Username,
#                 "PermitId": PermitId,
#                 "JobId": JobId,
#                 "RefId": RefId,
#                 "MsgId": MsgId,
#                 "AccountId": AccountId,
#                 "LoginStatus": head.get("LoginStatus", ""),
#                 "DateLastUpdated": str(head.get("DateLastUpdated", "")),
#                 "MailBoxId": head.get("MailBoxId", ""),
#                 "SeqPool": head.get("SeqPool", ""),
#                 "StartSequence": head.get("StartSequence", ""),
#                 "TradeNetMailboxID": head.get("TradeNetMailboxID", ""),
#                 "DeclarantName": head.get("DeclarantName", ""),
#                 "DeclarantCode": head.get("DeclarantCode", ""),
#                 "DeclarantTel": head.get("DeclarantTel", ""),
#                 "CRUEI": head.get("CRUEI", ""),
#                 "Code": head.get("Code", ""),
#                 "name": head.get("name", ""),
#                 "name1": head.get("name1", ""),
#                 "PermitNumber": "",
#                 "prmtStatus": "NEW",
#                 "CurrentDate": currentDate
#             })

#         except Exception as e:
#             print("--- DATABASE/LOGIC ERROR ---")
#             traceback.print_exc()
#             return Response({"error": str(e)}, status=500)

# New Permit
class CooNewPermit(APIView):
    print('hello')
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

            # 1. PER-USER count -> PermitId / RefId
            count_rows = SqlDb.execute_query(
                """
                SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                FROM CommonHeaderTbl
                WHERE PermitId LIKE %s
                """,
                [f"{Username}{refDate}%"]
            )
            count = count_rows[0]['Count'] if count_rows else 1

            RefId    = f"{count:03d}"
            PermitId = f"{Username}{refDate}{RefId}"

            # 2. GLOBAL count (all users, today) -> JobId / MsgId
            job_date = datetime.now().strftime("%y%m%d")  # matches K260610xxxxx format

            global_count_rows = SqlDb.execute_query(
                """
                SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                FROM CommonHeaderTbl
                WHERE JobId LIKE %s
                """,
                [f"K{job_date}%"]
            )
            global_count = global_count_rows[0]['Count'] if global_count_rows else 1

            JobId = f"K{job_date}{global_count:05d}"
            MsgId = f"{refDate}{global_count:04d}"

            print("PermitId:", PermitId, "| JobId:", JobId, "| MsgId:", MsgId,
                  "| count:", count, "| global_count:", global_count)

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


# class CopyCoo(APIView):
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
#                         VALUES (%s, 'COODEC', %s, %s, %s, %s)
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


# class CopyCoo(APIView):
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
#                         VALUES (%s, 'COODEC', %s, %s, %s, %s)
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

class CopyCoo(APIView):
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

                    # ── 2. Mirror the new row into COHeaderTbl ──
                    # Fields with no COHeaderTbl counterpart (CargoPackType, Inward*,
                    # License, ImporterCompanyCode, ClaimantPartyCode, EndUserCode, HBL,
                    # ArrivalDate/Time, LoadingPortCode, Voyage/Vessel/OBL/Conveyance
                    # inbound fields, ReleaseLocation, RecepitLocation, StorageLocation,
                    # BlanketStartDate, DepartureTime, OutMasterAirwayBill, VesselType,
                    # VesselNetRegTon, VesselNationality, Towing*, NextPort, LastPort,
                    # Total$ amounts, ResLoaName/RepLocName/RecepitLocName, outHAWB/INHAWB,
                    # seastore, Defrentprinting, Cnb, CondColor, gstVerified, Message,
                    # HandlingAgentCode, CustomerRemarks) are simply not mirrored.
                    cursor.execute("""
                        INSERT INTO COHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, ApplicationType, PreviousPermitNo,
                            OutwardTransportMode, ReferenceDocuments, COType,
                            CerDtlType1, CerDtlCopy1, CerDtlType2, CerDtlCopy2,
                            CurrencyCode, AdditionalCer, TransportDtl,
                            DeclarantCompanyCode, ExporterCompanyCode,
                            OutwardCarrierAgentCode, FreightForwarderCode,
                            CONSIGNEECode, Manufacturer,
                            DepartureDate, DischargePort, FinalDestinationCountry,
                            OutVoyageNumber, OutVesselName, OutConveyanceRefNo,
                            OutTransportId, OutFlightNO, OutAircraftRegNo,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            DeclareIndicator, NumberOfItems,
                            InternalRemarks, TradeRemarks,
                            Status, TouchUser, TouchTime, prmtStatus, PermitNumber,
                            EntryYear, Percomwealth, Gpsdonorcountry, Additionalrecieptant,
                            GrossReference, DeclarningFor, CertificateNumber,
                            MRDate, MRTime, TransmitId
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit,
                            OutwardTransportMode, ReferenceDocuments, COType,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            CurrencyCode, AddCerDtl, TransDtl,
                            DeclarantCompanyCode, ExporterCompanyCode,
                            OutwardCarrierAgentCode, FreightForwarderCode,
                            CONSIGNEECode, Manufacturer,
                            DepartureDate, DischargePort, FinalDestinationCountry,
                            OutVoyageNumber, OutVesselName, OutConveyanceRefNo,
                            OutTransportId, OutFlightNO, OutAircraftRegNo,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            DeclareIndicator, NumberOfItems,
                            InternalRemarks, TradeRemarks,
                            Status, TouchUser, TouchTime, prmtStatus, PermitNumber,
                            Entryyear, PerCommon, GSPDonorCountry, Recipient,
                            GrossReference, DeclarningFor, CertificateNumber,
                            MRDate, MRTime, TransmitId
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Item — copy Common(OLD) -> Common(NEW), unchanged ──
                    common_item_cols = [
                        "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                        "Contry", "EndUserDescription", "Brand", "Model",
                        "InHAWBOBL", "OutHAWBOBL", "DutiableQty", "DutiableUOM",
                        "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                        "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                        "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                        "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                        "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM", "ImPQty", "ImPUOM",
                        "PreferentialCode", "GSTRate", "GSTUOM", "GSTAmount",
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
                    ]
                    col_list = ", ".join(common_item_cols)
                    try:
                        cursor.execute(
                            f"""INSERT INTO CommonItemDtl (PermitId, {col_list})
                                SELECT %s, {col_list} FROM CommonItemDtl WHERE PermitId = %s""",
                            [new_permit_id, permit_id]
                        )
                    except Exception as err:
                        print(f"Warning copying CommonItemDtl: {err}")

                    # ── 4. Mirror Common(NEW) -> COItemDtl (renamed columns) ──
                    try:
                        cursor.execute("""
                            INSERT INTO COItemDtl (
                                PermitId, ItemNo, MessageType, HSCode, Description, Contry,
                                UnitPrice, UnitPriceCurrency, ExchangeRate, SumExchangeRate,
                                TotalLineAmount, CIFFOB, InvoiceQty, HSQTY, HSUOM,
                                ShippingMark, CerItemQty, CerItemUOM, ManfCostDate,
                                TextileCat, TextileQuotaQty, TextileQuotaQtyUOM,
                                ItemValue, InvoiceNumber, InvoiceDate, HSOnCer,
                                OriginCriterion, PerOrgainCRI, CertificateDes,
                                Touch_user, TouchTime, Hawblno
                            )
                            SELECT
                                %s, ItemNo, MessageType, HSCode, Description, Contry,
                                UnitPrice, UnitPriceCurrency, ExchangeRate, SumExchangeRate,
                                TotalLineAmount, CIFFOB, InvoiceQuantity, HSQty, HSUOM,
                                ShippingMarks1, CerItemQty, CerItemUOM, ManufactureCostDate,
                                TexCat, TexQuotaQty, TexQuotaUOM,
                                CIFValOfCer, CerInvNo, CerInvDate, HSCodeCer,
                                PreferentialCode, PerContent, CertificateDescription,
                                TouchUser, TouchTime, OutHAWBOBL
                            FROM CommonItemDtl WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonItemDtl -> COItemDtl: {err}")

                    # ── 5. File — COFileUpload has lowercase 'filepath', no 'Type' column ──
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
                            INSERT INTO COFileUpload (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filepath, Size)
                            SELECT %s, Sno, Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning mirroring CommonFile -> COFileUpload: {err}")

                    # ── 6. PMT (kept as-is — same schema as before) ──
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
                        VALUES (%s, 'COODEC', %s, %s, %s, %s)
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