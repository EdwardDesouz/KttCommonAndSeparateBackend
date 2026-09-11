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



class InnonpaymentList(APIView):
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

            MailBoxId = account_rows[0]["MailBoxId"]  # ← KEY CHANGE

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
                    CONVERT(varchar, t1.ArrivalDate, 105) AS ETA,
                    t1.PermitNumber AS PERMITNO,
                    i.Name + ' ' + i.Name1 AS IMPORTER,
                    t1.INHAWB AS HAWB,
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
                    AND t1.MessageType = 'INPDEC'
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId])

            else:
                nowdate = datetime.now() - timedelta(days=90)
                date_filter = nowdate.strftime("%Y/%m/%d")

                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'INPDEC'
                    AND CONVERT(varchar, t1.TouchTime, 111) >= %s
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId, date_filter])

            return Response(result)

        except Exception as e:
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)


# # New Permit
# class InnonpaymentNewPermit(APIView):
#     def get(self, request):
#         try:
#             Username = request.query_params.get("user")
#             if not Username:
#                 Username = request.session.get("Username")
#             if not Username:
#                 return Response({"error": "Session expired or User not provided"}, status=401)
           
#             refDate = datetime.now().strftime("%Y%m%d")
#             yy_mmdd = datetime.now().strftime("%Y%m%d")
#             currentDate = datetime.now().strftime("%d/%m/%Y")  

#             q_account = "SELECT AccountId FROM ManageUser WHERE UserName = %s"

#             account_rows = SqlDb.execute_query(q_account, [Username])
#             if not account_rows:
#                 return Response({"error": "User not found"}, status=404)
#             AccountId = account_rows[0]['AccountId']

#             # q_ref = """
#             #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
#             #     FROM CommonHeaderTbl 
#             #     WHERE MSGId LIKE %s AND MessageType = 'INPDEC'
#             # """
#             # ref_rows = SqlDb.execute_query(q_ref, [f"%{refDate}%"])
#             # ref_rows = SqlDb.execute_query(q_ref, [f"{refDate}%"])

#             # ref_rows = SqlDb.execute_query(
#             #     """
#             #     SELECT ISNULL(COUNT(*), 0) + 1 AS Count
#             #     FROM CommonHeaderTbl
#             #     WHERE PermitId LIKE %s
#             #     """,
#             #     [f"{Username}{refDate}%"]
#             # )
           
           
           
#             # RefId = "%03d" % (ref_rows[0]['Count'] if ref_rows else 1)

#             # q_job = """
#             #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
#             #     FROM PermitCount 
#             #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'INPDEC'
#             # """

#             # job_rows = SqlDb.execute_query(
#                 # """
#                 # # SELECT ISNULL(MAX(CAST(SUBSTRING(JobId, 2, 6) + RIGHT(JobId, 5) AS BIGINT)), 0) + 1 AS Count
#                 # # FROM CommonHeaderTbl
#                 # """
#             # job_rows = SqlDb.execute_query(
#             #     """
#             #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
#             #     FROM PermitCount 
#             #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'INPDEC'
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
class InnonpaymentNewPermit(APIView):
    def get(self, request):
        try:
            Username = request.query_params.get("user")
            if not Username:
                Username = request.session.get("Username")
            if not Username:
                return Response({"error": "Session expired or User not provided"}, status=401)

            refDate = datetime.now().strftime("%Y%m%d")       # 20260911 -> RefId/PermitId/MsgId ku
            yy_mmdd = datetime.now().strftime("%y%m%d")        # 260911   -> JobId ku
            jobDate = datetime.now().strftime("%Y-%m-%d")       # 2026-09-11 -> PermitCount.TouchTime match ku
            currentDate = datetime.now().strftime("%d/%m/%Y")

            # ── 1. AccountId ─────────────────────────────────────────────
            q_account = "SELECT AccountId FROM ManageUser WHERE UserName = %s"
            account_rows = SqlDb.execute_query(q_account, [Username])
            if not account_rows:
                return Response({"error": "User not found"}, status=404)
            AccountId = account_rows[0]['AccountId']

            # ── 2. RefId — CommonHeaderTbl vechi, Username+refDate scoped ──
            ref_rows = SqlDb.execute_query(
                """
                SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                FROM CommonHeaderTbl
                WHERE PermitId LIKE %s AND MessageType = 'INPDEC'
                """,
                [f"{Username}{refDate}%"]
            )
            RefId = "%03d" % (ref_rows[0]['Count'] if ref_rows else 1)
            PermitId = f"{Username}{refDate}{RefId}"

            # ── 3. JobIdCount — PermitCount vechi, AccountId scoped (OLD LOGIC) ──
            job_rows = SqlDb.execute_query(
                """
                SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                FROM PermitCount
                WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'INPDEC'
                """,
                [f"%{jobDate}%", AccountId]
            )
            JobIdCount = job_rows[0]['Count'] if job_rows else 1

            # ── 4. JobId + MsgId — SAME JobIdCount, different format (OLD LOGIC) ──
            JobId = f"K{yy_mmdd}{'%05d' % JobIdCount}"
            MsgId = f"{refDate}{'%04d' % JobIdCount}"

            print("PermitId:", PermitId, "| JobId:", JobId, "| MsgId:", MsgId,
                  "| RefId:", RefId, "| JobIdCount:", JobIdCount, "| AccountId:", AccountId)

            # ── 5. Declarant/company profile info ────────────────────────
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



# class CopyInnonpayment(APIView):
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


#                     ref_count = cursor.fetchone()[0]

#                     ref_id = f"{ref_count:03d}"
#                     # JobId + MsgId — scoped to AccountId + today in PermitCount
#                     cursor.execute("""
#                         SELECT ISNULL(COUNT(*), 0) + 1
#                         FROM PermitCount
#                         WHERE TouchTime LIKE %s AND AccountId = %s
#                     """, [f"%{today_dash}%", account_id])
#                     job_count = cursor.fetchone()[0]
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
#                             NULL, NULL,
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
#                         VALUES (%s, 'INPDEC', %s, %s, %s, %s)
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


# class CopyInnonpayment(APIView):
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

#                     ref_id = f"{ref_count:03d}"
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
#                             NULL, NULL,
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
#                     cursor.execute("""
#                         INSERT INTO InnonHeaderTbl (
#                             Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
#                             MessageType, DeclarationType, PreviousPermit, CargoPackType,
#                             InwardTransportMode, BGIndicator, SupplyIndicator,
#                             ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
#                             ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
#                             ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
#                             VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
#                             TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
#                             ReleaseLocation, ReleaseLocName, RecepitLocation,
#                             TotalOuterPack, TotalOuterPackUOM,
#                             TotalGrossWeight, TotalGrossWeightUOM,
#                             GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
#                             DeclareIndicator, NumberOfItems,
#                             TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
#                             TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
#                             Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
#                             RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
#                             CondColor, TransmitId, gstVerified
#                         )
#                         SELECT
#                             Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
#                             MessageType, DeclarationType, PreviousPermit, CargoPackType,
#                             InwardTransportMode, BGIndicator, SupplyIndicator,
#                             ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
#                             ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
#                             ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
#                             VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
#                             TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
#                             ReleaseLocation, ResLoaName, RecepitLocation,
#                             TotalOuterPack, TotalOuterPackUOM,
#                             TotalGrossWeight, TotalGrossWeightUOM,
#                             GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
#                             DeclareIndicator, NumberOfItems,
#                             TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
#                             TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
#                             Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
#                             RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
#                             CondColor, TransmitId, gstVerified
#                         FROM CommonHeaderTbl
#                         WHERE PermitId = %s
#                     """, [new_permit_id])

#                     child_tables = {
#                         "CommonInvoiceDtl":  ("InnoninvoiceDtl",[
#                             "SNo", "InvoiceNo", "InvoiceDate", "TermType",
#                             "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
#                             "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
#                             "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
#                             "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
#                             "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
#                             "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
#                             "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
#                         ]),
#                         "CommonItemDtl": ("InNonItemDtl",[
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
#                         ]),
#                         "CommonCASCDtl":("INNONCASCDtl", [
#                                 "ItemNo", "ProductCode", "Quantity", "ProductUOM",
#                                 "RowNo", "CascCode1", "CascCode2", "CascCode3",
#                                 "MessageType", "TouchUser", "TouchTime", "EndUserDes", "CASCId"
#                             ]),
#                         "CommonCPCDtl":("InNonCPCDtl", [
#                             "MessageType", "RowNo", "CPCType",
#                             "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
#                             "TouchUser", "TouchTime"
#                         ]),
#                         "CommonContainerDtl":("InnonContainerDtl", [
#                             "RowNo", "ContainerNo", "Size", "Weight", "SealNo", "MessageType","TouchUser", "TouchTime"
#                         ]),
#                     }


#                     for src_table, (dst_table, cols) in child_tables.items():
#                         col_list = ", ".join(cols)
#                         # Copy Common -> Common (from OLD permit, as before)
#                         try:
#                             cursor.execute(
#                                 f"""INSERT INTO {src_table} (PermitId, {col_list})
#                                     SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
#                                 [new_permit_id, permit_id]
#                             )
#                         except Exception as err:
#                             print(f"Warning copying {src_table}: {err}")

#                         # Copy Common(NEW) -> In-payment target — single read of what we
#                         # just wrote, so both tables end up byte-identical for this permit.
#                         try:
#                             cursor.execute(
#                                 f"""INSERT INTO {dst_table} (PermitId, {col_list})
#                                     SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
#                                 [new_permit_id, new_permit_id]
#                             )
#                         except Exception as err:
#                             print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

#                     # File & PMT — column names differ, handle explicitly
#                     try:
#                         cursor.execute("""
#                             INSERT INTO CommonFile (PermitId, Sno,Name, ContentType, Data,
#                                 DocumentType, TouchUser, TouchTime, filePath, Size, Type)
#                             SELECT %s, Sno,Name, ContentType, Data, DocumentType,
#                                 TouchUser, TouchTime, filePath, Size, Type
#                             FROM CommonFile WHERE PermitId = %s
#                         """, [new_permit_id, permit_id])
#                         cursor.execute("""
#                             INSERT INTO InNonFile (PermitId,Sno, Name, ContentType, Data,
#                                 DocumentType, TouchUser, TouchTime, filePath, Size, Type)
#                                 SELECT %s, Sno,Name, ContentType, Data, DocumentType,
#                                 TouchUser, TouchTime, filePath, Size, Type
#                             FROM CommonFile WHERE PermitId = %s
#                         """, [new_permit_id, new_permit_id])
#                     except Exception as err:
#                         print(f"Warning copying File: {err}")

#                     try:
#                         cursor.execute("""
#                             INSERT INTO CommonPMT (PermitId, ConditionCode, ConditionDesc, PermitNumber)
#                             SELECT %s, ConditionCode, ConditionDesc, PermitNumber
#                             FROM CommonPMT WHERE PermitId = %s
#                         """, [new_permit_id, permit_id])
#                         # InPMT has a different schema (Sno, CertificateNumber, StartDate, etc.)
#                         # — usually populated on approval, not at copy time. Skip unless you
#                         # need to carry PMT conditions forward too; adjust columns if so.
#                     except Exception as err:
#                         print(f"Warning copying PMT: {err}")

#                     cursor.execute("""
#                         INSERT INTO PermitCount
#                             (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
#                         VALUES (%s, 'INPDEC', %s, %s, %s, %s)
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


class CopyInnonpayment(APIView):
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
            ref_date   = now.strftime("%Y%m%d")
            job_date   = now.strftime("%y%m%d")
            today_dash = now.strftime("%Y-%m-%d")

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

                    ref_id = f"{ref_count:03d}"
                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    # ── 1. Insert new CommonHeaderTbl row (source = OLD permit) ──
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
                            NULL, NULL,
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

                    # ── 2. Mirror the NEW CommonHeaderTbl row into InnonHeaderTbl ──
                    # Reads from the NEW row (new_permit_id) — single source of truth,
                    # no ID recomputation needed.
                    cursor.execute("""
                        INSERT INTO InnonHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, Inhabl, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ReleaseLocaName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepilocaName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified,
                            FinalDestinationCountry
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ResLoaName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified,
                            ISNULL(FinalDestinationCountry, '')   -- adjust default if column has a specific fallback code
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: Common(OLD) -> Common(NEW), then
                    #       Common(NEW) -> InNon* target (mirrored) ──
                    #
                    # Format:
                    #   "CommonTable": ("InNonTable", [src_cols], [dst_cols])
                    #   -> use this 4-tuple form whenever column names differ
                    #      between Common* and InNon* (e.g. CommonItemDtl)
                    #
                    #   "CommonTable": ("InNonTable", [cols])
                    #   -> use this 3-tuple shorthand when column names are
                    #      IDENTICAL on both sides
                    child_tables = {
                        "CommonInvoiceDtl": ("InNonInvoiceDtl", [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                        "CommonItemDtl": (
                            "InNonItemDtl",
                            [  # SOURCE (CommonItemDtl) column names
                                "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                                "Contry", "Brand", "Model",
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
                                "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                                "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                                "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                            ],
                            [  # TARGET (InNonItemDtl) column names — 3 renamed vs source
                                "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                                "Contry", "Brand", "Model",
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
                                "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                                "Enginecapacity", "Optioncahrge", "OptionalSumtotal",
                                "OptionalSumExchage", "Engineuom", "Orginregdate"
                            ]
                        ),
                        "CommonCASCDtl": ("INNONCASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
                            "RowNo", "CascCode1", "CascCode2", "CascCode3",
                            "MessageType", "TouchUser", "TouchTime", "CASCId"
                        ]),
                        "CommonCPCDtl": ("InNonCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("InnonContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, mapping in child_tables.items():
                        if len(mapping) == 2:
                            dst_table, src_cols = mapping
                            dst_cols = src_cols  # identical column names both sides
                        else:
                            dst_table, src_cols, dst_cols = mapping  # ItemDtl: names differ

                        src_col_list = ", ".join(src_cols)
                        dst_col_list = ", ".join(dst_cols)

                        # Copy Common -> Common (from OLD permit, as before)
                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {src_col_list})
                                    SELECT %s, {src_col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")

                        # Copy Common(NEW) -> InNon target — single read of what we
                        # just wrote, so both tables end up in sync for this permit.
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {dst_col_list})
                                    SELECT %s, {src_col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── File & PMT — column names differ, handle explicitly ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno,Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                        cursor.execute("""
                            INSERT INTO InNonFile (PermitId,Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, Size, Type)
                                SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning copying File: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO CommonPMT (PermitId, ConditionCode, ConditionDesc, PermitNumber)
                            SELECT %s, ConditionCode, ConditionDesc, PermitNumber
                            FROM CommonPMT WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                        # InnonPMT has a different schema (Sno, CertificateNumber, StartDate, etc.)
                        # — usually populated on approval, not at copy time. Skip unless you
                        # need to carry PMT conditions forward too; adjust columns if so.
                    except Exception as err:
                        print(f"Warning copying PMT: {err}")

                    # ── PermitCount — MessageType = 'INPDEC' for non-payment ──
                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'INPDEC', %s, %s, %s, %s)
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

class TransmitInpayment(APIView):

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

                    print(f"[TransmitInpayment] Source {permit_id}: "
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
                            'IPTDEC', %s,
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
                            print(f"[TransmitInpayment] {table}: {rows_inserted} row(s) "
                                  f"copied from {permit_id} -> {new_permit_id}")

                            if rows_inserted == 0 and table in ("CommonInvoiceDtl", "CommonItemDtl"):
                                warnings.append(
                                    f"{permit_id}: 0 rows copied into {table} "
                                    f"(source had {src_invoice_count if table=='CommonInvoiceDtl' else src_item_count} rows)"
                                )

                            if "MessageType" in cols:
                                cursor.execute(
                                    f"UPDATE {table} SET MessageType = 'IPTDEC' WHERE PermitId = %s",
                                    [new_permit_id]
                                )

                        except Exception as err:
                            print(f"[TransmitInpayment] ERROR copying {table}: {err}")
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
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
                        """,
                        [new_permit_id, account_id, msg_id, username, touch_time]
                    )

                    copied_permits.append(new_permit_id)

                    # Increment counters for next permit in this batch
                    job_count += 1
                    ref_count += 1

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) transmitted successfully as InPayment",
                "copiedPermits": copied_permits,
                "declarationType": declaration_type,
                "messageType": "IPTDEC",
                "warnings": warnings,
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)



class PostInnonAmendTable(APIView):
    table = "CommonAmend"
    in_table = "InNonAmend"
    allowed_columns = {
        "Permitno", "AmendmentCount", "UpdateIndicator",
        "ReplacementPermitno", "DescriptionOfReason",
        "PermitExtension", "ExtendImportPeriod",
        "DeclarationIndigator", "AmendType",
        "TouchUser", "TouchTime", "MSGId",
    }

    # CommonAmend column -> InNonAmend column (only TouchTime differs)
    IN_TABLE_COLUMN_MAP = {
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
        permit_id = item.get("Permitno", "")

        if not msg_id:
            return Response({"error": "MSGId is required"}, status=400)

        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns provided"}, status=400)

        try:
            # ── 1. CommonAmend upsert ──
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

            # ── 2. Mirror into InNonAmend (remapped TouchTime -> TouchTme) ──
            in_item = {}
            for src_col, dst_col in self.IN_TABLE_COLUMN_MAP.items():
                if src_col in item:
                    in_item[dst_col] = item.get(src_col)

            if in_item:
                existing_in = SqlDb.execute_query(
                    f"SELECT Id FROM {self.in_table} WHERE MSGId = %s", [msg_id]
                )
                if existing_in:
                    in_update_cols = [c for c in in_item.keys() if c != "MSGId"]
                    if in_update_cols:
                        in_set_clause = ", ".join([f"{c} = %s" for c in in_update_cols])
                        in_values = [in_item[c] for c in in_update_cols] + [msg_id]
                        SqlDb.execute_query(
                            f"UPDATE {self.in_table} SET {in_set_clause} WHERE MSGId = %s",
                            in_values
                        )
                else:
                    in_col_str = ", ".join(in_item.keys())
                    in_ph_str = ", ".join(["%s"] * len(in_item))
                    in_values = list(in_item.values())
                    SqlDb.execute_query(
                        f"INSERT INTO {self.in_table} ({in_col_str}) VALUES ({in_ph_str})",
                        in_values
                    )

            SqlDb.commit()
            # prmtStatus intentionally left untouched — it's already 'AMD'
            # from CopyAmend at the point the amend copy was created,
            # matching the inpayment PostAmendTable pattern.

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response(
            {"Result": f"Amend record {action} successfully", "MSGId": msg_id},
            status=201
        )

class CopyInnonAmend(APIView):
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
                    #      force MessageType to INPDEC and prmtStatus to AMD ──
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
                            'INPDEC', DeclarationType, PreviousPermit, CargoPackType,
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

                    # ── 2. Mirror into InnonHeaderTbl (column set/renames match
                    #      CopyInnonpayment's mirror step) ──
                    cursor.execute("""
                        INSERT INTO InnonHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode,ConsigneeCode, Inhabl, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ReleaseLocaName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepilocaName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified,
                            FinalDestinationCountry
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode,CONSIGNEECode, HBL, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ResLoaName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified,
                            ISNULL(FinalDestinationCountry, '')
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: Common(OLD) -> Common(NEW), then
                    #       Common(NEW) -> InNon* target (mirrored) ──
                    child_tables = {
                        "CommonInvoiceDtl": ("InNonInvoiceDtl", [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                        "CommonItemDtl": (
                            "InNonItemDtl",
                            [  # SOURCE (CommonItemDtl) column names
                                "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                                "Contry", "Brand", "Model",
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
                                "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                                "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                                "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                            ],
                            [  # TARGET (InNonItemDtl) column names — 3 renamed vs source
                                "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                                "Contry", "Brand", "Model",
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
                                "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                                "Enginecapacity", "Optioncahrge", "OptionalSumtotal",
                                "OptionalSumExchage", "Engineuom", "Orginregdate"
                            ]
                        ),
                        "CommonCASCDtl": ("INNONCASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
                            "RowNo", "CascCode1", "CascCode2", "CascCode3",
                            "MessageType", "TouchUser", "TouchTime", "CASCId"
                        ]),
                        "CommonCPCDtl": ("InNonCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("InnonContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, mapping in child_tables.items():
                        if len(mapping) == 2:
                            dst_table, src_cols = mapping
                            dst_cols = src_cols
                        else:
                            dst_table, src_cols, dst_cols = mapping

                        src_col_list = ", ".join(src_cols)
                        dst_col_list = ", ".join(dst_cols)

                        # Common -> Common (from OLD permit)
                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {src_col_list})
                                    SELECT %s, {src_col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")

                        # Common(NEW) -> InNon target — read of what we just wrote
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {dst_col_list})
                                    SELECT %s, {src_col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── 4. File & PMT — column names/targets differ, handled explicitly ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno,Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                        cursor.execute("""
                            INSERT INTO InNonFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, Size, Type)
                                SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning copying File: {err}")

                    try:
                        cursor.execute("""
                            INSERT INTO CommonPMT (PermitId, ConditionCode, ConditionDesc, PermitNumber)
                            SELECT %s, ConditionCode, ConditionDesc, PermitNumber
                            FROM CommonPMT WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                        # InnonPMT has a different schema, populated on approval —
                        # skip here, same as CopyInnonpayment.
                    except Exception as err:
                        print(f"Warning copying PMT: {err}")

                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'INPDEC', %s, %s, %s, %s)
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





# Cancel
class PostInnonCancelTable(APIView):
    table = "CommonCancel"
    in_table = "InNonCancel"
    allowed_columns = {
        "Permitno", "UpdateIndicator", "ReplacementPermitno", "ReasonForCancel",
        "DescriptionOfReason", "DeclarationIndigator",
        "TouchUser", "TouchTime", "MSGId", "CancelType"
    }

    # CommonCancel column -> InNonCancel column
    # (InNonCancel has typo'd column names: ResonForCancel, TouchTme)
    IN_TABLE_COLUMN_MAP = {
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

            # ── 2. Mirror into InNonCancel (remapped column names) ──
            in_item = {}
            for src_col, dst_col in self.IN_TABLE_COLUMN_MAP.items():
                if src_col in item:
                    in_item[dst_col] = item.get(src_col)

            if in_item:
                existing_in = SqlDb.execute_query(
                    f"SELECT Id FROM {self.in_table} WHERE MSGId = %s", [msg_id]
                )
                if existing_in:
                    in_update_cols = [c for c in in_item.keys() if c != "MSGId"]
                    if in_update_cols:
                        in_set_clause = ", ".join([f"{c} = %s" for c in in_update_cols])
                        in_values = [in_item[c] for c in in_update_cols] + [msg_id]
                        SqlDb.execute_query(
                            f"UPDATE {self.in_table} SET {in_set_clause} WHERE MSGId = %s",
                            in_values
                        )
                else:
                    in_col_str = ", ".join(in_item.keys())
                    in_ph_str = ", ".join(["%s"] * len(in_item))
                    in_values = list(in_item.values())
                    SqlDb.execute_query(
                        f"INSERT INTO {self.in_table} ({in_col_str}) VALUES ({in_ph_str})",
                        in_values
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
                        "UPDATE InnonHeaderTbl SET prmtStatus = 'CNL' WHERE PermitId = %s",
                        [permit_id]
                    )
                except Exception as err:
                    print(f"Warning updating InnonHeaderTbl prmtStatus: {err}")
                SqlDb.commit()

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response({"Result": f"Cancel record {action} successfully", "MSGId": msg_id}, status=201)

# copyinnoncancel

class CopyInnonCancel(APIView):
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
                    #      force MessageType to INPDEC and prmtStatus to CNL ──
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
                            'INPDEC', DeclarationType, PreviousPermit, CargoPackType,
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

                    # ── 2. Mirror into InnonHeaderTbl (same shape as CopyInnonAmend) ──
                    cursor.execute("""
                        INSERT INTO InnonHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode,ConsigneeCode, Inhabl, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ReleaseLocaName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepilocaName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified,
                            FinalDestinationCountry
                        )
                        SELECT
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode,CONSIGNEECode, HBL, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ResLoaName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified,
                            ISNULL(FinalDestinationCountry, '')
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: Common(OLD) -> Common(NEW), then
                    #       Common(NEW) -> InNon* target (mirrored) ──
                    child_tables = {
                        "CommonInvoiceDtl": ("InNonInvoiceDtl", [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                        "CommonItemDtl": (
                            "InNonItemDtl",
                            [
                                "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                                "Contry", "Brand", "Model",
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
                                "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                                "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                                "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                            ],
                            [
                                "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                                "Contry", "Brand", "Model",
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
                                "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                                "Enginecapacity", "Optioncahrge", "OptionalSumtotal",
                                "OptionalSumExchage", "Engineuom", "Orginregdate"
                            ]
                        ),
                        "CommonCASCDtl": ("INNONCASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
                            "RowNo", "CascCode1", "CascCode2", "CascCode3",
                            "MessageType", "TouchUser", "TouchTime", "CASCId"
                        ]),
                        "CommonCPCDtl": ("InNonCPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("InnonContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, mapping in child_tables.items():
                        if len(mapping) == 2:
                            dst_table, src_cols = mapping
                            dst_cols = src_cols
                        else:
                            dst_table, src_cols, dst_cols = mapping

                        src_col_list = ", ".join(src_cols)
                        dst_col_list = ", ".join(dst_cols)

                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {src_col_list})
                                    SELECT %s, {src_col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")

                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {dst_col_list})
                                    SELECT %s, {src_col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── 4. File & PMT ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno,Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                        cursor.execute("""
                            INSERT INTO InNonFile (PermitId, Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, Size, Type)
                                SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, new_permit_id])
                    except Exception as err:
                        print(f"Warning copying File: {err}")

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
                        VALUES (%s, 'INPDEC', %s, %s, %s, %s)
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
