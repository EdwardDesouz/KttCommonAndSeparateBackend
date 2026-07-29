from django.db import connections,connection
from django.contrib.sessions.models import Session
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
from xml.etree.ElementTree import Element, SubElement, tostring, fromstring, register_namespace
from decimal import Decimal
import xlwt
import pandas as pd
import zipfile
import logging
import json
import traceback
import os
import io
import requests



"""The below sqldb is used for commanly we can use everywhere if we need
 database's table just create new class and use inherit"""
class SqlDb:
    database_name = "default"
    @classmethod
    def execute_query(cls, query, params=None):
        """---Class method to execute SELECT queries and return list of dicts---"""
        with connections[cls.database_name].cursor() as cursor:
            cursor.execute(query, params or [])
            if cursor.description:
                columns = [col[0] for col in cursor.description]
                return [dict(zip(columns, row)) for row in cursor.fetchall()]
            return cursor.rowcount
    @classmethod
    def commit(cls):
        connections[cls.database_name].commit()

#18-02-26
#--------Login Page-----

class LoginAPIView(APIView):
    def post(self,request):
        Username = request.data.get("Username")
        password = request.data.get("Password")
        check = request.data.get("ChkLogin")


        if not Username or not password:
            print("Error: Missing credentials")
            return Response({"error": "Username and Password are required"}, status=status.HTTP_400_BAD_REQUEST)

        query = "SELECT * FROM ManageUser WHERE UserName = %s"
        users = SqlDb.execute_query(query, [Username])

        if not users:
            print(f"Error: User {Username} not found in database")
            return Response({"error": "Enter Correct Password and Username"}, status=status.HTTP_401_UNAUTHORIZED)

        user = users[0]

        if user["Password"] != password:
            print("Error: Password mismatch")
            return Response({"error": "Enter Correct Password and Username"}, status=status.HTTP_401_UNAUTHORIZED)

        if user["LoginStatus"] == "True":
            
            if check:
              
                request.session["Username"] = Username
                request.session["Permit_Id"] = "NEW"
                request.session.modified = True
                request.session.save()

                self._invalidate_other_sessions(Username, keep_session_key=request.session.session_key)
                return Response({ 
                    "success": "Success",
                    "AccountId": user["AccountId"],
                    "Username": Username  
                })
            else:            
                return Response({"error": "User Already Login with another System"}, status=status.HTTP_403_FORBIDDEN)
        
        update_query = "UPDATE ManageUser SET LoginStatus='True' WHERE UserName=%s"
        SqlDb.execute_query(update_query, [Username])
        SqlDb.commit()

        request.session["Username"] = Username
        request.session["Permit_Id"] = "NEW"
        request.session.modified = True
        request.session.save()

        return Response(
            {"success": "Success",
            "AccountId": user["AccountId"],
            "Username": Username },
        )

    def _invalidate_other_sessions(self, username, keep_session_key):
        for s in Session.objects.all():
            try:
                data = s.get_decoded()
            except Exception:
                continue 
            if data.get("Username") == username and s.session_key != keep_session_key:
                s.delete()

# Logout
class LogoutAPIView(APIView):
    def post(self, request):
        Username = request.data.get("Username")
        if not Username:
            return Response({"error": "Username is required"}, status=400)
        SqlDb.execute_query(
            "UPDATE ManageUser SET LoginStatus='False' WHERE UserName=%s",
            [Username]
        )
        SqlDb.commit()
        try:
            request.session.flush()
        except Exception:
            pass
        return Response({"success": "Logged out successfully"})


"""The below code has used for getting all datas from table using restapi format"""
#05-02-26
#-------Header Table-------
class GetCommonHeaderTable(APIView):
    table = "CommonHeaderTbl"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)


class GetCommonHeaderByPermitId(APIView):
    table = "CommonHeaderTbl"
    def get(self, request):
            try:
                permit_id = request.GET.get("PermitId")
                if not permit_id:
                    return Response(
                        {"error": "PermitId is required"},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
                rows = SqlDb.execute_query(query, [permit_id])
                if not rows:
                    return Response(
                        {"error": "Record not found"},
                        status=status.HTTP_404_NOT_FOUND
                    )
                data = rows[0] if isinstance(rows, list) else rows
                return Response(data, status=status.HTTP_200_OK)
            except Exception as e:
                return Response(
                    {"error": str(e)},
                    status=status.HTTP_500_INTERNAL_SERVER_ERROR
                )


# class PostCommonHeaderTable(APIView):
#     table = "CommonHeaderTbl"
#     allowed_columns = {
#         "Refid", "JobId", "MSGId", "PermitId", "TradeNetMailboxID", "MessageType",
#         "DeclarationType", "PreviousPermit", "CargoPackType", "InwardTransportMode",
#         "OutwardTransportMode", "BGIndicator", "SupplyIndicator", "ReferenceDocuments",
#         "License", "COType", "Entryyear", "GSPDonorCountry", "CerDetailtype1",
#         "CerDetailCopies1", "CerDetailtype2", "CerDetailCopies2", "PerCommon",
#         "CurrencyCode", "AddCerDtl", "TransDtl", "Recipient", "DeclarantCompanyCode",
#         "ExporterCompanyCode", "Inwardcarriercode", "OutwardCarrierAgentCode",
#         "FreightForwarderCode", "ImporterCompanyCode", "InwardCarrierAgentCode",
#         "CONSIGNEECode", "ClaimantPartyCode", "EndUserCode", "Manufacturer",
#         "HBL", "ArrivalDate", "ArrivalTime", "LoadingPortCode", "VoyageNumber",
#         "VesselName", "OceanBillofLadingNo", "ConveyanceRefNo", "TransportId",
#         "FlightNO", "AircraftRegNo", "MasterAirwayBill", "ReleaseLocation",
#         "RecepitLocation", "StorageLocation", "BlanketStartDate", "DepartureDate",
#         "DepartureTime", "DischargePort", "FinalDestinationCountry", "OutVoyageNumber",
#         "OutVesselName", "OutOceanBillofLadingNo", "VesselType", "VesselNetRegTon",
#         "VesselNationality", "TowingVesselID", "TowingVesselName", "NextPort",
#         "LastPort", "OutConveyanceRefNo", "OutTransportId", "OutFlightNO",
#         "OutAircraftRegNo", "OutMasterAirwayBill", "TotalOuterPack",
#         "TotalOuterPackUOM", "TotalGrossWeight", "TotalGrossWeightUOM",
#         "GrossReference", "TradeRemarks", "InternalRemarks", "DeclareIndicator",
#         "NumberOfItems", "TotalCIFFOBValue", "TotalGSTTaxAmt", "TotalExDutyAmt",
#         "TotalCusDutyAmt", "TotalODutyAmt", "TotalAmtPay", "Status",
#         "TouchUser", "TouchTime", "PermitNumber", "prmtStatus", "ResLoaName",
#         "RepLocName", "RecepitLocName", "outHAWB", "INHAWB", "seastore",
#         "CertificateNumber", "Defrentprinting", "Cnb", "DeclarningFor",
#         "MRDate", "MRTime", "CondColor", "TransmitId", "gstVerified","Message","HandlingAgentCode"
#     }
#     def post(self, request):
#         payloads = request.data
#         if not payloads:
#             return Response({"error": "No data provided"}, status=400)

#         if not isinstance(payloads, list):
#             payloads = [payloads]

#         inserted_count = 0
#         JobId = ""
#         MsgId = ""

#         try:
#             for item in payloads:
#                 PermitId = item.get("PermitId", "")
#                 TouchUser = item.get("TouchUser", "")
#                 TouchTime = item.get("TouchTime", "")
#                 jobDate = datetime.now().strftime("%Y-%m-%d")
#                 refDate  = datetime.now().strftime("%Y%m%d")

#                 # Get AccountId
#                 account_rows = SqlDb.execute_query(
#                     "SELECT AccountId FROM ManageUser WHERE UserName = %s", [TouchUser]
#                 )
#                 if not account_rows:
#                     return Response({"error": f"User '{TouchUser}' not found"}, status=404)
#                 AccountId = account_rows[0]['AccountId']

#                 # Check if this PermitId already saved (EDIT)
#                 existing_pcount = SqlDb.execute_query(
#                     "SELECT MsgId FROM PermitCount WHERE PermitId = %s", [PermitId]
#                 )

#                 if existing_pcount:
#                     # EDIT — reuse existing MsgId
#                     MsgId = existing_pcount[0]['MsgId']
#                     existing_header = SqlDb.execute_query(
#                         f"SELECT JobId FROM {self.table} WHERE PermitId = %s", [PermitId]
#                     )
#                     JobId = existing_header[0]['JobId'] if existing_header else item.get("JobId", "")

#                 else:
#                     # NEW — count from CommonHeaderTbl (already committed, reliable)
#                     # 6-6-2026
#                     # count_rows = SqlDb.execute_query(
#                     #     """SELECT ISNULL(COUNT(*), 0) + 1 AS Count
#                     #     FROM PermitCount
#                     #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = %s""",
#                     #     [f"%{jobDate}%", AccountId, item.get("MessageType", "IPTDEC")]
#                     # )
#                     # end
#                     count_rows = SqlDb.execute_query(
#                         """
#                         SELECT ISNULL(COUNT(*), 0) + 1 AS Count
#                         FROM CommonHeaderTbl
#                         WHERE PermitId LIKE %s
#                         """,
#                         [f"{TouchUser}{refDate}%"]
#                     )

#                     count = count_rows[0]['Count'] if count_rows else 1

#                     # JobId = f"K{datetime.now().strftime('%y%m%d')}{'%05d' % count}"
#                     JobId = f"K{datetime.now().strftime('%y%m%d')}{count:05d}"
#                     # MsgId = f"{datetime.now().strftime('%Y%m%d')}{'%04d' % count}"
#                     MsgId = f"{refDate}{count:04d}"


#                     # Try insert PermitCount, skip if duplicate
#                     try:
#                         SqlDb.execute_query(
#                             """INSERT INTO PermitCount 
#                             (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
#                             VALUES (%s, %s, %s, %s, %s, %s)""",
#                             [PermitId, item.get("MessageType", ""), AccountId, MsgId, TouchUser, TouchTime]
#                         )
#                         SqlDb.commit()
#                     except Exception:
#                         # PK duplicate — row already exists, fetch the real MsgId
#                         SqlDb.execute_query("SELECT 1")  # reset cursor state
#                         existing_fallback = SqlDb.execute_query(
#                             "SELECT MsgId FROM PermitCount WHERE PermitId = %s", [PermitId]
#                         )
#                         if existing_fallback:
#                             MsgId = existing_fallback[0]['MsgId']
#                             existing_header = SqlDb.execute_query(
#                                 f"SELECT JobId FROM {self.table} WHERE PermitId = %s", [PermitId]
#                             )
#                             JobId = existing_header[0]['JobId'] if existing_header else JobId

#                 # Set correct values
#                 item["JobId"] = JobId
#                 item["MSGId"] = MsgId

#                 # If TotalGSTTaxAmt > 10000, force Status to 'WFA'
#                 try:
#                     gst_amt = float(item.get("TotalGSTTaxAmt") or 0)
#                 except (ValueError, TypeError):
#                     gst_amt = 0.0

#                 if gst_amt > 10000:
#                     item["Status"] = "WFA"

#                 # Sanitize
#                 # sanitized_item = {}
#                 # for col in self.allowed_columns:
#                 #     val = item.get(col)
#                 #     sanitized_item[col] = val if val is not None else ""

#                 sanitized_item = {}
#                 for col in self.allowed_columns:
#                     val = item.get(col)
#                     if val == "" or val is None:
#                         sanitized_item[col] = None
#                     else:
#                         sanitized_item[col] = val

#                 print("sanitized_item:", sanitized_item)

#                 # Upsert CommonHeaderTbl
#                 existing_header_check = SqlDb.execute_query(
#                     f"SELECT 1 FROM {self.table} WHERE PermitId = %s", [PermitId]
#                 )
#                 if existing_header_check:
#                     set_clause = ", ".join([f"{col} = %s" for col in sanitized_item.keys()])
#                     values = list(sanitized_item.values()) + [PermitId]
#                     SqlDb.execute_query(
#                         f"UPDATE {self.table} SET {set_clause} WHERE PermitId = %s", values
#                     )
#                 else:
#                     columns = ", ".join(sanitized_item.keys())
#                     placeholders = ", ".join(["%s"] * len(sanitized_item))
#                     SqlDb.execute_query(
#                         f"INSERT INTO {self.table} ({columns}) VALUES ({placeholders})",
#                         list(sanitized_item.values())
#                     )

#                 SqlDb.commit()
#                 inserted_count += 1

#             return Response(
#                 {
#                     "message": f"{inserted_count} record(s) saved successfully",
#                     "JobId": JobId,
#                     "MSGId": MsgId,
#                 },
#                 status=201
#             )
#         except Exception as e:
#             import traceback
#             traceback.print_exc()
#             return Response({"error": f"Database Error: {str(e)}"}, status=400)

class PostCommonHeaderTable(APIView):
    table = "CommonHeaderTbl"
    allowed_columns = {
        "Refid", "JobId", "MSGId", "PermitId", "TradeNetMailboxID", "MessageType",
        "DeclarationType", "PreviousPermit", "CargoPackType", "InwardTransportMode",
        "OutwardTransportMode", "BGIndicator", "SupplyIndicator", "ReferenceDocuments",
        "License", "COType", "Entryyear", "GSPDonorCountry", "CerDetailtype1",
        "CerDetailCopies1", "CerDetailtype2", "CerDetailCopies2", "PerCommon",
        "CurrencyCode", "AddCerDtl", "TransDtl", "Recipient", "DeclarantCompanyCode",
        "ExporterCompanyCode", "Inwardcarriercode", "OutwardCarrierAgentCode",
        "FreightForwarderCode", "ImporterCompanyCode", "InwardCarrierAgentCode",
        "CONSIGNEECode", "ClaimantPartyCode", "EndUserCode", "Manufacturer",
        "HBL", "ArrivalDate", "ArrivalTime", "LoadingPortCode", "VoyageNumber",
        "VesselName", "OceanBillofLadingNo", "ConveyanceRefNo", "TransportId",
        "FlightNO", "AircraftRegNo", "MasterAirwayBill", "ReleaseLocation",
        "RecepitLocation", "StorageLocation", "BlanketStartDate", "DepartureDate",
        "DepartureTime", "DischargePort", "FinalDestinationCountry", "OutVoyageNumber",
        "OutVesselName", "OutOceanBillofLadingNo", "VesselType", "VesselNetRegTon",
        "VesselNationality", "TowingVesselID", "TowingVesselName", "NextPort",
        "LastPort", "OutConveyanceRefNo", "OutTransportId", "OutFlightNO",
        "OutAircraftRegNo", "OutMasterAirwayBill", "TotalOuterPack",
        "TotalOuterPackUOM", "TotalGrossWeight", "TotalGrossWeightUOM",
        "GrossReference", "TradeRemarks", "InternalRemarks", "DeclareIndicator",
        "NumberOfItems", "TotalCIFFOBValue", "TotalGSTTaxAmt", "TotalExDutyAmt",
        "TotalCusDutyAmt", "TotalODutyAmt", "TotalAmtPay", "Status",
        "TouchUser", "TouchTime", "PermitNumber", "prmtStatus", "ResLoaName",
        "RepLocName", "RecepitLocName", "outHAWB", "INHAWB", "seastore",
        "CertificateNumber", "Defrentprinting", "Cnb", "DeclarningFor",
        "MRDate", "MRTime", "CondColor", "TransmitId", "gstVerified","Message","HandlingAgentCode","CustomerRemarks"
    }
    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)

        if not isinstance(payloads, list):
            payloads = [payloads]

        inserted_count = 0
        JobId = ""
        MsgId = ""

        try:
            for item in payloads:
                PermitId = item.get("PermitId", "")
                TouchUser = item.get("TouchUser", "")
                TouchTime = item.get("TouchTime", "")
                jobDate = datetime.now().strftime("%Y-%m-%d")
                refDate  = datetime.now().strftime("%Y%m%d")
                job_date = datetime.now().strftime("%y%m%d")

                # Get AccountId
                account_rows = SqlDb.execute_query(
                    "SELECT AccountId FROM ManageUser WHERE UserName = %s", [TouchUser]
                )
                if not account_rows:
                    return Response({"error": f"User '{TouchUser}' not found"}, status=404)
                AccountId = account_rows[0]['AccountId']

                # Check if this PermitId already saved (EDIT)
                existing_pcount = SqlDb.execute_query(
                    "SELECT MsgId FROM PermitCount WHERE PermitId = %s", [PermitId]
                )

                if existing_pcount:
                    # EDIT — reuse existing MsgId
                    MsgId = existing_pcount[0]['MsgId']
                    existing_header = SqlDb.execute_query(
                        f"SELECT JobId FROM {self.table} WHERE PermitId = %s", [PermitId]
                    )
                    JobId = existing_header[0]['JobId'] if existing_header else item.get("JobId", "")

                else:
                    # NEW — GLOBAL count from CommonHeaderTbl (all users, today)
                    count_rows = SqlDb.execute_query(
                        """
                        SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                        FROM CommonHeaderTbl
                        WHERE JobId LIKE %s
                        """,
                        [f"K{job_date}%"]
                    )

                    count = count_rows[0]['Count'] if count_rows else 1

                    JobId = f"K{job_date}{count:05d}"
                    MsgId = f"{refDate}{count:04d}"

                    # Try insert PermitCount, skip if duplicate
                    try:
                        SqlDb.execute_query(
                            """INSERT INTO PermitCount 
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                            VALUES (%s, %s, %s, %s, %s, %s)""",
                            [PermitId, item.get("MessageType", ""), AccountId, MsgId, TouchUser, TouchTime]
                        )
                        SqlDb.commit()
                    except Exception:
                        # PK duplicate — row already exists, fetch the real MsgId
                        SqlDb.execute_query("SELECT 1")  # reset cursor state
                        existing_fallback = SqlDb.execute_query(
                            "SELECT MsgId FROM PermitCount WHERE PermitId = %s", [PermitId]
                        )
                        if existing_fallback:
                            MsgId = existing_fallback[0]['MsgId']
                            existing_header = SqlDb.execute_query(
                                f"SELECT JobId FROM {self.table} WHERE PermitId = %s", [PermitId]
                            )
                            JobId = existing_header[0]['JobId'] if existing_header else JobId

                # Set correct values
                item["JobId"] = JobId
                item["MSGId"] = MsgId

                # If TotalGSTTaxAmt > 10000, force Status to 'WFA'
                message_type = (item.get("MessageType") or "").strip().upper()
                incoming_status = (item.get("Status") or "").strip().upper()

                

                try:
                    gst_amt = float(item.get("TotalGSTTaxAmt") or 0)
                except (ValueError, TypeError):
                    gst_amt = 0.0

                NON_SUBMIT_STATUSES = {"SAVEASDRF", "DRF", "DISCONNECT"}

                if incoming_status in NON_SUBMIT_STATUSES:
                    item["Status"] = incoming_status
                elif message_type in ("IPTDEC", "INPDEC") and gst_amt > 10000:
                    item["Status"] = "WFA"
                else:
                    item["Status"] = "NEW"

                sanitized_item = {}
                for col in self.allowed_columns:
                    val = item.get(col)
                    if val == "" or val is None:
                        sanitized_item[col] = None
                    else:
                        sanitized_item[col] = val

                print("sanitized_item:", sanitized_item)

                # Upsert CommonHeaderTbl
                existing_header_check = SqlDb.execute_query(
                    f"SELECT 1 FROM {self.table} WHERE PermitId = %s", [PermitId]
                )
                if existing_header_check:
                    set_clause = ", ".join([f"{col} = %s" for col in sanitized_item.keys()])
                    values = list(sanitized_item.values()) + [PermitId]
                    SqlDb.execute_query(
                        f"UPDATE {self.table} SET {set_clause} WHERE PermitId = %s", values
                    )
                else:
                    columns = ", ".join(sanitized_item.keys())
                    placeholders = ", ".join(["%s"] * len(sanitized_item))
                    SqlDb.execute_query(
                        f"INSERT INTO {self.table} ({columns}) VALUES ({placeholders})",
                        list(sanitized_item.values())
                    )

                SqlDb.commit()
                inserted_count += 1

            return Response(
                {
                    "message": f"{inserted_count} record(s) saved successfully",
                    "JobId": JobId,
                    "MSGId": MsgId,
                },
                status=201
            )
        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": f"Database Error: {str(e)}"}, status=400)


class EditCommonHeaderByPermit(APIView):
    table = "CommonHeaderTbl"
    def put(self, request, permit_id):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id", "JobId", "MSGId", "PermitId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(permit_id)
        query = f"UPDATE {self.table} SET {set_clause} WHERE PermitId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with PermitId {permit_id} updated successfully"})

class DeletePermit(APIView):
    table = "CommonHeaderTbl"
    def get(self, request):
        try:
            permit_id = request.GET.get("PermitId")
            if not permit_id:
                return Response(
                    {"error": "PermitId is required"},
                    status=status.HTTP_400_BAD_REQUEST
                )
            update_query = f"UPDATE {self.table} SET Status = 'DEL' WHERE PermitId = %s"
            result = SqlDb.execute_query(update_query, [permit_id])
            SqlDb.commit()
            if not result:
                return Response(
                    {"error": "Record not found"},
                    status=status.HTTP_404_NOT_FOUND
                )
            return Response(
                {"message": f"Record with PermitId {permit_id} marked as DEL successfully"}
            )
        except Exception as e:
            return Response({"error": str(e)}, status=400)

#---------InvoiceTable---------

class CommonInvoiceTable(APIView):
    table="CommonInvoiceDtl"
    def get(self,request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)


class GetInvoiceByInvoiceNo(APIView):
    table = "CommonInvoiceDtl"
    def get(self, request, invoice_no, permit_id):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE InvoiceNo = %s AND PermitId = %s"
        data = SqlDb.execute_query(query, [invoice_no, permit_id])
        if not data:
            return Response(
                {
                    "message": f"No records found for InvoiceNo {invoice_no} and PermitId {permit_id}"
                },
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)

class GetInvoiceByEditPermitId(APIView):
    table = "CommonInvoiceDtl"

    def get(self, request):
            try:
                permit_id = request.GET.get("PermitId")
                if not permit_id:
                    return Response(
                        {"error": "PermitId is required"},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
                rows = SqlDb.execute_query(query, [permit_id])
                if not rows:
                    return Response(
                        {"error": "Record not found"},
                        status=status.HTTP_404_NOT_FOUND
                    )
                
                return Response(rows, status=status.HTTP_200_OK)
            except Exception as e:
                return Response(
                    {"error": str(e)},
                    status=status.HTTP_500_INTERNAL_SERVER_ERROR
                )


class GetInvoiceByPermitId(APIView):
    table = "CommonInvoiceDtl"

    def get(self, request, permit_id):

        query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
        data = SqlDb.execute_query(query, [permit_id])

        if not data:
            # return Response(
            #     {"message": f"No invoices found for PermitId {permit_id}"},
            #     status=status.HTTP_404_NOT_FOUND
            # )
            return Response([], status=200)
        float_fields = [
            "TIExRate", "OTCExRate", "OTCAmount", "OTCSAmount",
            "FCExRate", "FCAmount", "FCSAmount", "ICExRate",
            "ICAmount", "ICSAmount", "OTCCharge", "FCCharge"
        ]

        for row in data:
            for field in float_fields:
                if field in row and row[field] is not None:
                    row[field] = str(row[field])

        return Response(data, status=status.HTTP_200_OK)


class DeleteInvoice(APIView):
    table = "CommonInvoiceDtl"
    def post(self, request):
        permit_id = request.data.get("PermitId")
        s_no = request.data.get("SNo")
        if not permit_id or not s_no:
            return Response(
                {"error": "PermitId and SNo required"},
                status=400
            )
        try:
            with connections['default'].cursor() as cursor:
                cursor.execute(
                    f"""
                    DELETE FROM {self.table}
                    WHERE PermitId = %s AND SNo = %s
                    """,
                    [permit_id, s_no]
                )
                cursor.execute(
                    f"""
                    UPDATE {self.table}
                    SET SNo = SNo - 1
                    WHERE PermitId = %s AND SNo > %s
                    """,
                    [permit_id, s_no]
                )
                cursor.execute(
                    f"""
                    SELECT *
                    FROM {self.table}
                    WHERE PermitId = %s
                    ORDER BY SNo
                    """,
                    [permit_id]
                )
                columns = [col[0] for col in cursor.description]
                records = [
                    dict(zip(columns, row))
                    for row in cursor.fetchall()
                ]
                connections['default'].commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({
            "Result": "Invoice deleted successfully",
            "Records": records
        })



# class DeleteInvoice(APIView):
#     table = "CommonInvoiceDtl"

#     def delete(self, request, invoice_no, permit_id):
#         invoice_no = invoice_no.strip()
#         permit_id = permit_id.strip()

#         select_query = f"""
#             SELECT 1
#             FROM {self.table}
#             WHERE InvoiceNo = %s AND PermitId = %s
#         """
#         existing = SqlDb.execute_query(select_query, [invoice_no, permit_id])

#         if not existing:
#             return Response(
#                 {"error": f"No record found with InvoiceNo {invoice_no} and PermitId {permit_id}"},
#                 status=status.HTTP_404_NOT_FOUND
#             )

#         delete_query = f"""
#             DELETE FROM {self.table}
#             WHERE InvoiceNo = %s AND PermitId = %s
#         """

#         try:
#             SqlDb.execute_query(delete_query, [invoice_no, permit_id])
#             SqlDb.commit()
#         except Exception as e:
#             import traceback
#             print(traceback.format_exc())
#             return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)

#         return Response(
#             {
#                 "message": "Record deleted successfully",
#                 "invoiceNo": invoice_no,
#                 "permitId": permit_id
#             },
#             status=status.HTTP_200_OK
#         )


class PostInvoiceTable(APIView):
    table = "CommonInvoiceDtl"
    allowed_columns = {
        "PermitId","SNo","InvoiceNo","InvoiceDate","TermType",
        "AdValoremIndicator","PreDutyRateIndicator",
        "SupplierImporterRelationship","SupplierCode","ImportPartyCode",
        "TICurrency","TIExRate","TIAmount","TISAmount",
        "OTCCharge","OTCCurrency","OTCExRate","OTCAmount","OTCSAmount",
        "FCCharge","FCCurrency","FCExRate","FCAmount","FCSAmount",
        "ICCharge","ICCurrency","ICExRate","ICAmount","ICSAmount",
        "CIFSUMAmount",
        "GSTPercentage","GSTSUMAmount",
        "MessageType","TouchUser","TouchTime","ChkOtherInv"
    }

    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)
        if not isinstance(payloads, list):
            payloads = [payloads]
        permit_id = None
        try:
            with connections['default'].cursor() as cursor:
                for item in payloads:
                    if not isinstance(item, dict):
                        item = dict(item)
                    item.pop("Id", None)
                    # Filter allowed columns
                    columns = [k for k in item.keys() if k in self.allowed_columns]
                    if not columns:
                        continue
                    permit_id = item.get("PermitId")
                    sno = item.get("SNo")
                    if not permit_id or not sno:
                        return Response(
                            {"error": "PermitId and SNo are required"},
                            status=400
                        )
                    # Check existing record
                    cursor.execute(
                        f"""
                        SELECT COUNT(*)
                        FROM {self.table}
                        WHERE PermitId=%s AND SNo=%s
                        """,
                        [permit_id, sno]
                    )
                    exists = cursor.fetchone()[0] > 0
                    if exists:
                        update_columns = [
                            col for col in columns
                            if col not in ["PermitId", "SNo"]
                        ]
                        set_clause = ", ".join(
                            [f"{col}=%s" for col in update_columns]
                        )
                        values = [item[col] for col in update_columns]
                        values += [permit_id, sno]
                        query = f"""
                        UPDATE {self.table}
                        SET {set_clause}
                        WHERE PermitId=%s AND SNo=%s
                        """
                        cursor.execute(query, values)
                        action = "updated"
                    else:
                        placeholders = ", ".join(["%s"] * len(columns))
                        values = [item[col] for col in columns]
                        query = f"""
                        INSERT INTO {self.table}
                        ({", ".join(columns)})
                        VALUES ({placeholders})
                        """
                        cursor.execute(query, values)
                        action = "inserted"
                connections['default'].commit()
        except Exception as e:
            return Response(
                {"error": f"Error saving record: {str(e)}"},
                status=400
            )
        # Fetch all invoice records for this permit
        try:
            with connections['default'].cursor() as cursor:
                fetch_columns = ", ".join(self.allowed_columns)
                cursor.execute(
                    f"""
                    SELECT {fetch_columns}
                    FROM {self.table}
                    WHERE PermitId=%s
                    ORDER BY SNo
                    """,
                    [permit_id]
                )
                rows = cursor.fetchall()
                records = [
                    dict(zip(self.allowed_columns, row))
                    for row in rows
                ]
        except Exception:
            records = []
        return Response(
            {
                "Result": f"Invoice {action} successfully",
                "Records": records
            },
            status=201
        )


class EditInvoiceByInvoiceNo(APIView):
    table = "CommonInvoiceDtl"
    def put(self, request, invoice_no, permit_id):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id", "InvoiceNo","PermitId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.extend([invoice_no, permit_id])
        query = f"UPDATE {self.table} SET {set_clause} WHERE InvoiceNo = %s AND PermitId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with InvoiceNo {invoice_no} PermitId {permit_id} updated successfully"})

#06-02-26
#---------item---------

class GetCommonItemTabel(APIView):
    table = "CommonItemDtl"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonItemByItemNo(APIView):
    table = "CommonItemDtl"
    def get(self, request, item_no, permit_id):
        # Parameterized query to prevent SQL injection
        query = f"""
            SELECT *
            FROM {self.table}
            WHERE ItemNo = %s AND PermitId = %s
        """
        data = SqlDb.execute_query(query, [item_no, permit_id])
        if not data:
            return Response(
                {
                    "message": f"No records found for ItemNo {item_no} and PermitId {permit_id}"
                },
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data, status=status.HTTP_200_OK)

class GetItemByEditPermitId(APIView):
    table = "CommonItemDtl"
    def get(self, request):
            try:
                permit_id = request.GET.get("PermitId")
                if not permit_id:
                    return Response(
                        {"error": "PermitId is required"},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                query = f"SELECT * FROM {self.table} WHERE PermitId = %s ORDER BY ItemNo"
                # query = f"""
                # SELECT i.*, h.InnonPayment
                # FROM {self.table} i
                # LEFT JOIN Hscode h ON h.HSCode = i.HSCode
                # WHERE i.PermitId = %s
                # """

                rows = SqlDb.execute_query(query, [permit_id])
                if not rows:
                    return Response(
                        {"error": "Record not found"},
                        status=status.HTTP_404_NOT_FOUND
                    )              
                return Response(rows, status=status.HTTP_200_OK)
            except Exception as e:
                return Response(
                    {"error": str(e)},
                    status=status.HTTP_500_INTERNAL_SERVER_ERROR
                )

# class DeleteItem(APIView):
#     table = "CommonItemDtl"
#     def delete(self, request, item_no, permit_id):
#         item_no = item_no.strip()
#         permit_id = permit_id.strip()
#         # Check if record exists
#         select_query = f"""
#             SELECT 1
#             FROM {self.table}
#             WHERE ItemNo = %s AND PermitId = %s
#         """
#         existing = SqlDb.execute_query(select_query, [item_no, permit_id])

#         if not existing:
#             return Response(
#                 {"error": f"No record found with ItemNo {item_no} and PermitId {permit_id}"},
#                 status=status.HTTP_404_NOT_FOUND
#             )

#         # Delete record
#         delete_query = f"""
#             DELETE FROM {self.table}
#             WHERE ItemNo = %s AND PermitId = %s
#         """

#         try:
#            SqlDb.execute_query(delete_query, [item_no, permit_id])
#            SqlDb.commit()
#         except Exception as e:
#             import traceback
#             print(traceback.format_exc())
#             return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)


#         return Response(
#             {
#                 "message": "Record deleted successfully",
#                 "item_no": item_no,
#                 "permitId": permit_id
#             },
#             status=status.HTTP_200_OK
#         )

# class DeleteItem(APIView):
#     table = "CommonItemDtl"
#     casc_table = "CommonCASCDtl"

#     def post(self, request):
#         permit_id = request.data.get("PermitId")
#         i_no = request.data.get("ItemNo")
#         if not permit_id or not i_no:
#             return Response(
#                 {"error": "PermitId and ItemNo required"},
#                 status=400
#             )
#         try:
#             with connections['default'].cursor() as cursor:
#                 cursor.execute(
#                     f"""
#                     DELETE FROM {self.table}
#                     WHERE PermitId = %s AND ItemNo = %s
#                     """,
#                     [permit_id, i_no]
#                 )
#                 cursor.execute(
#                     f"DELETE FROM {self.casc_table} WHERE PermitId=%s AND ItemNo=%s",
#                     [permit_id, i_no]
#                 )
#                 cursor.execute(
#                     f"""
#                     UPDATE {self.table}
#                     SET ItemNo = ItemNo - 1
#                     WHERE PermitId = %s AND ItemNo > %s
#                     """,
#                     [permit_id, i_no]
#                 )
#                 cursor.execute(
#                     f"UPDATE {self.casc_table} SET ItemNo = ItemNo - 1 WHERE PermitId=%s AND ItemNo > %s",
#                     [permit_id, i_no]
#                 )
#                 cursor.execute(
#                     f"""
#                     SELECT *
#                     FROM {self.table}
#                     WHERE PermitId = %s
#                     ORDER BY ItemNo
#                     """,
#                     [permit_id]
#                 )
#                 columns = [col[0] for col in cursor.description]
#                 records = [
#                     dict(zip(columns, row))
#                     for row in cursor.fetchall()
#                 ]
#                 connections['default'].commit()
#         except Exception as e:
#             return Response({"error": str(e)}, status=400)
#         return Response({
#             "Result": "Item deleted successfully",
#             "Records": records
#         })

# class DeleteItem(APIView):
#     table = "CommonItemDtl"
#     casc_table = "CommonCASCDtl"

#     def post(self, request):
#         permit_id = request.data.get("PermitId")
#         item_nos = request.data.get("ItemNos")  # list of ints
#         if not permit_id or not item_nos:
#             return Response({"error": "PermitId and ItemNos required"}, status=400)
#         try:
#             with connections['default'].cursor() as cursor:
#                 placeholders = ",".join(["%s"] * len(item_nos))

#                 # 1. Delete CASC rows for all selected items in one shot
#                 cursor.execute(
#                     f"DELETE FROM {self.casc_table} WHERE PermitId=%s AND ItemNo IN ({placeholders})",
#                     [permit_id] + item_nos
#                 )
#                 # 2. Delete the Item rows in one shot
#                 cursor.execute(
#                     f"DELETE FROM {self.table} WHERE PermitId=%s AND ItemNo IN ({placeholders})",
#                     [permit_id] + item_nos
#                 )

#                 # 3. Figure out the new sequential numbering for what's left
#                 cursor.execute(
#                     f"SELECT ItemNo FROM {self.table} WHERE PermitId=%s ORDER BY ItemNo",
#                     [permit_id]
#                 )
#                 remaining = [row[0] for row in cursor.fetchall()]
#                 mapping = {old: new for new, old in enumerate(remaining, start=1)}
#                 changed = {old: new for old, new in mapping.items() if old != new}

#                 # 4. Renumber only the rows that actually shifted
#                 #    (negative staging avoids unique-constraint collisions mid-renumber)
#                 for old in changed:
#                     cursor.execute(
#                         f"UPDATE {self.table} SET ItemNo=%s WHERE PermitId=%s AND ItemNo=%s",
#                         [-old, permit_id, old]
#                     )
#                     cursor.execute(
#                         f"UPDATE {self.casc_table} SET ItemNo=%s WHERE PermitId=%s AND ItemNo=%s",
#                         [-old, permit_id, old]
#                     )
#                 for old, new in changed.items():
#                     cursor.execute(
#                         f"UPDATE {self.table} SET ItemNo=%s WHERE PermitId=%s AND ItemNo=%s",
#                         [new, permit_id, -old]
#                     )
#                     cursor.execute(
#                         f"UPDATE {self.casc_table} SET ItemNo=%s WHERE PermitId=%s AND ItemNo=%s",
#                         [new, permit_id, -old]
#                     )

#                 connections['default'].commit()

#                 # 5. One final fetch, not one per item
#                 cursor.execute(
#                     f"SELECT * FROM {self.table} WHERE PermitId=%s ORDER BY ItemNo",
#                     [permit_id]
#                 )
#                 columns = [col[0] for col in cursor.description]
#                 records = [dict(zip(columns, row)) for row in cursor.fetchall()]
#         except Exception as e:
#             return Response({"error": str(e)}, status=400)

#         return Response({
#             "Result": "Items deleted successfully",
#             "Records": records
#         })

class DeleteItem(APIView):
    table = "CommonItemDtl"
    casc_table = "CommonCASCDtl"

    def post(self, request):
        permit_id = request.data.get("PermitId")
        item_nos = request.data.get("ItemNos")
        if not permit_id or not item_nos:
            return Response({"error": "PermitId and ItemNos required"}, status=400)
        try:
            item_nos = [int(i) for i in item_nos]
            with connections['default'].cursor() as cursor:
                placeholders = ",".join(["%s"] * len(item_nos))

                cursor.execute(
                    f"DELETE FROM {self.casc_table} WHERE PermitId=%s AND ItemNo IN ({placeholders})",
                    [permit_id] + item_nos
                )
                cursor.execute(
                    f"DELETE FROM {self.table} WHERE PermitId=%s AND ItemNo IN ({placeholders})",
                    [permit_id] + item_nos
                )

                cursor.execute(
                    f"SELECT ItemNo FROM {self.table} WHERE PermitId=%s ORDER BY ItemNo",
                    [permit_id]
                )
                remaining = [row[0] for row in cursor.fetchall()]
                mapping = {old: new for new, old in enumerate(remaining, start=1)}
                changed = {old: new for old, new in mapping.items() if old != new}

                # Single CASE-based UPDATE per table — no per-row round trips,
                # so 1 deleted item and 40 deleted items cost the same.
                if changed:
                    case_sql = " ".join(
                        f"WHEN ItemNo = {old} THEN {new}" for old, new in changed.items()
                    )
                    id_list = ",".join(str(o) for o in changed.keys())

                    cursor.execute(
                        f"""
                        UPDATE {self.table}
                        SET ItemNo = CASE {case_sql} END
                        WHERE PermitId = %s AND ItemNo IN ({id_list})
                        """,
                        [permit_id]
                    )
                    cursor.execute(
                        f"""
                        UPDATE {self.casc_table}
                        SET ItemNo = CASE {case_sql} END
                        WHERE PermitId = %s AND ItemNo IN ({id_list})
                        """,
                        [permit_id]
                    )

                connections['default'].commit()

                cursor.execute(
                    f"SELECT * FROM {self.table} WHERE PermitId=%s ORDER BY ItemNo",
                    [permit_id]
                )
                columns = [col[0] for col in cursor.description]
                records = [dict(zip(columns, row)) for row in cursor.fetchall()]
        except Exception as e:
            return Response({"error": str(e)}, status=400)

        return Response({
            "Result": "Items deleted successfully",
            "Records": records
        })



# class PostItemTable(APIView):
#     table ="CommonItemDtl"
#     allowed_columns = {"ItemNo","PermitId","MessageType","HSCode","Description","DGIndicator","Contry",
# "EndUserDescription","Brand","Model","InHAWBOBL","OutHAWBOBL","DutiableQty","DutiableUOM","TotalDutiableQty",
# "TotalDutiableUOM","InvoiceQuantity","HSQty","HSUOM","AlcoholPer","InvoiceNo","ChkUnitPrice","UnitPrice",
# "UnitPriceCurrency","ExchangeRate","SumExchangeRate","TotalLineAmount","InvoiceCharges","CIFFOB","OPQty",
# "OPUOM","IPQty","IPUOM","InPqty","InPUOM","ImPQty","ImPUOM","PreferentialCode","GSTRate","GSTUOM","GSTAmount",
# "ExciseDutyRate","ExciseDutyUOM","ExciseDutyAmount","CustomsDutyRate","CustomsDutyUOM","CustomsDutyAmount","OtherTaxRate",
# "OtherTaxUOM","OtherTaxAmount","CurrentLot","PreviousLot","LSPValue","Making","ShippingMarks1","ShippingMarks2","ShippingMarks3",
# "ShippingMarks4","CerItemQty","CerItemUOM","CIFValOfCer",
# "ManufactureCostDate","TexCat","TexQuotaQty","TexQuotaUOM","CerInvNo","CerInvDate","OriginOfCer","HSCodeCer","PerContent","CertificateDescription",
# "TouchUser","TouchTime","VehicleType","OptionalChrgeUOM","EngineCapcity","Optioncahrge","OptionalSumtotal",
# "OptionalSumExchage","EngineCapUOM","orignaldatereg"}
#     def post(self, request):
#         payloads = request.data
#         if not payloads:
#             return Response({"error": "No data provided"}, status=400)
#         if not isinstance(payloads, list):
#             payloads = [payloads]  # handle single-object case
#         inserted_count = 0
#         for item in payloads:
#             if not isinstance(item, dict):
#                 item = dict(item)  # convert QueryDict / other mapping to dict
#             item.pop("Id", None)  # safe removal of Id
#             columns = ", ".join(item.keys())
#             placeholders = ", ".join(["%s"] * len(item))
#             values = list(item.values())
#             query = f"""INSERT INTO {self.table} ({columns}) VALUES ({placeholders})"""
#             #print("query:",query)
#             try:
#                 SqlDb.execute_query(query, values)
#                 inserted_count += 1
#             except Exception as e:
#                 return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
#         SqlDb.commit()
#         return Response(
#             {"message": f"{inserted_count} record(s) inserted successfully"},
#             status=201
#         )

class PostItemTable(APIView):
    table = "CommonItemDtl"
    allowed_columns = [
        "ItemNo","PermitId","MessageType","HSCode","Description","DGIndicator","Contry",
        "EndUserDescription","Brand","Model","InHAWBOBL","OutHAWBOBL","DutiableQty","DutiableUOM","TotalDutiableQty",
        "TotalDutiableUOM","InvoiceQuantity","HSQty","HSUOM","AlcoholPer","InvoiceNo","ChkUnitPrice","UnitPrice",
        "UnitPriceCurrency","ExchangeRate","SumExchangeRate","TotalLineAmount","InvoiceCharges","CIFFOB","OPQty",
        "OPUOM","IPQty","IPUOM","InPqty","InPUOM","ImPQty","ImPUOM","PreferentialCode","GSTRate","GSTUOM","GSTAmount",
        "ExciseDutyRate","ExciseDutyUOM","ExciseDutyAmount","CustomsDutyRate","CustomsDutyUOM","CustomsDutyAmount","OtherTaxRate",
        "OtherTaxUOM","OtherTaxAmount","CurrentLot","PreviousLot","LSPValue","Making","ShippingMarks1","ShippingMarks2","ShippingMarks3",
        "ShippingMarks4","CerItemQty","CerItemUOM","CIFValOfCer",
        "ManufactureCostDate","TexCat","TexQuotaQty","TexQuotaUOM","CerInvNo","CerInvDate","OriginOfCer","HSCodeCer","PerContent","CertificateDescription",
        "TouchUser","TouchTime","VehicleType","OptionalChrgeUOM","EngineCapcity","Optioncahrge","OptionalSumtotal",
        "OptionalSumExchage","EngineCapUOM","orignaldatereg"
    ]

    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)
        if not isinstance(payloads, list):
            payloads = [payloads]
        permit_id = None
        action = ""
        try:
            with connections['default'].cursor() as cursor:
                for item in payloads:
                    if not isinstance(item, dict):
                        item = dict(item)
                    item.pop("Id", None)
                    columns = [k for k in item.keys() if k in self.allowed_columns]
                    if not columns:
                        continue
                    permit_id = item.get("PermitId")
                    item_no = item.get("ItemNo")
                    if not permit_id or not item_no:
                        return Response(
                            {"error": "PermitId and ItemNo are required"},
                            status=400
                        )
                    cursor.execute(
                        f"""
                        SELECT COUNT(*)
                        FROM {self.table}
                        WHERE PermitId=%s AND ItemNo=%s
                        """,
                        [permit_id, item_no]
                    )
                    exists = cursor.fetchone()[0] > 0
                    if exists:
                        update_columns = [
                            col for col in columns
                            if col not in ["PermitId", "ItemNo"]
                        ]
                        if update_columns:
                            set_clause = ", ".join(
                                [f"{col}=%s" for col in update_columns]
                            )
                            values = [item[col] for col in update_columns]
                            values += [permit_id, item_no]
                            query = f"""
                            UPDATE {self.table}
                            SET {set_clause}
                            WHERE PermitId=%s AND ItemNo=%s
                            """
                            cursor.execute(query, values)
                        action = "updated"
                    else:
                        placeholders = ", ".join(["%s"] * len(columns))
                        values = [item[col] for col in columns]
                        query = f"""
                        INSERT INTO {self.table}
                        ({", ".join(columns)})
                        VALUES ({placeholders})
                        """
                        cursor.execute(query, values)
                        action = "inserted"

                connections['default'].commit()

        except Exception as e:
            return Response(
                {"error": f"Error saving record: {str(e)}"},
                status=400
            )
        try:
            with connections['default'].cursor() as cursor:
                fetch_columns = ", ".join(self.allowed_columns)

                cursor.execute(
                    f"""
                    SELECT {fetch_columns}
                    FROM {self.table}
                    WHERE PermitId=%s
                    ORDER BY ItemNo
                    """,
                    [permit_id]
                )

                rows = cursor.fetchall()

                records = [
                    dict(zip(self.allowed_columns, row))
                    for row in rows
                ]

        except Exception:
            records = []

        return Response(
            {
                "Result": f"Item {action} successfully",
                "Records": records
            },
            status=201
        )



class EditItemByItemNo(APIView):
    table = "CommonItemDtl"
    def put(self, request, item_no, permit_id):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        # Do not allow updating key fields
        for key in ["Id", "ItemNo", "PermitId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.extend([item_no, permit_id])
        query = f"UPDATE {self.table} SET {set_clause} WHERE ItemNo = %s AND PermitId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with ItemNo {item_no} PermitId {permit_id} updated successfully"})

#---------casc-------

class GetCommonCascTabel(APIView):
    table = "CommonCASCDtl"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonCascByPermitId(APIView):
    table = "CommonCASCDtl"
    def get(self, request, permit_id):
        query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
        data = SqlDb.execute_query(query, [permit_id])
        if not data:
            return Response(
                {"message": f"No records found for PermitId {permit_id}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteCasc(APIView):
    table = "CommonCASCDtl"
    def post(self, request):
        permit_id = request.data.get("PermitId")
        i_no = request.data.get("ItemNo")
        if not permit_id or not i_no:
            return Response(
                {"error": "PermitId and ItemNo required"},
                status=400
            )
        try:
            with connections['default'].cursor() as cursor:
                cursor.execute(
                    f"""
                    DELETE FROM {self.table}
                    WHERE PermitId = %s AND ItemNo = %s
                    """,
                    [permit_id, i_no]
                )
                cursor.execute(
                    f"""
                    UPDATE {self.table}
                    SET ItemNo = ItemNo - 1
                    WHERE PermitId = %s AND ItemNo > %s
                    """,
                    [permit_id, i_no]
                )
                cursor.execute(
                    f"""
                    SELECT *
                    FROM {self.table}
                    WHERE PermitId = %s
                    ORDER BY ItemNo
                    """,
                    [permit_id]
                )
                columns = [col[0] for col in cursor.description]
                records = [
                    dict(zip(columns, row))
                    for row in cursor.fetchall()
                ]
                connections['default'].commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({
            "Result": "CASC deleted successfully",
            "Records": records
        })

class DeleteCascByCascId(APIView):
    table = "CommonCASCDtl"
    def delete(self, request, casc_id, row_no, permit_id):
        try:
            with connections['default'].cursor() as cursor:
                cursor.execute(
                    f"""
                    DELETE FROM {self.table}
                    WHERE CASCId = %s
                    AND RowNo = %s
                    AND PermitId = %s
                    """,
                    [casc_id, row_no, permit_id]
                )
                rows_affected = cursor.rowcount
            if rows_affected == 0:
                return Response(
                    {"message": "No CASC row found to delete"},
                    status=404
                )
            return Response(
                {"message": "CASC row deleted successfully"},
                status=200
            )
        except Exception as e:
            return Response({"error": str(e)}, status=400)

# class PostCascTable(APIView):
#     table = "CommonCASCDtl"
#     allowed_columns = [
#         "ItemNo", "ProductCode", "Quantity", "ProductUOM", "RowNo",
#         "CascCode1", "CascCode2", "CascCode3",
#         "PermitId", "MessageType",
#         "TouchUser", "TouchTime",
#         "CASCId", "EndUserDes"
#     ]
#     def post(self, request):
#         payloads = request.data
#         if not payloads:
#             return Response({"error": "No data provided"}, status=400)
#         if not isinstance(payloads, list):
#             payloads = [payloads]
#         inserted_count = 0
#         updated_count = 0
#         try:
#             with connections['default'].cursor() as cursor:
#                 for item in payloads:
#                     if not isinstance(item, dict):
#                         item = dict(item)
#                     item.pop("Id", None)
#                     casc_id = item.get("CASCId")
#                     permit_id = item.get("PermitId")
#                     row_no = item.get("RowNo")
#                     if not casc_id or not permit_id or row_no is None:
#                         continue
#                     cursor.execute(
#                         f"""
#                         SELECT COUNT(*)
#                         FROM {self.table}
#                         WHERE CASCId=%s
#                           AND PermitId=%s
#                           AND RowNo=%s
#                         """,
#                         [casc_id, permit_id, row_no]
#                     )
#                     exists = cursor.fetchone()[0] > 0
#                     if exists:

#                         update_columns = [
#                             col for col in self.allowed_columns
#                             if col not in ["CASCId", "PermitId", "RowNo"]
#                             and col in item
#                         ]
#                         if update_columns:

#                             set_clause = ", ".join([f"{col}=%s" for col in update_columns])

#                             values = [item.get(col) for col in update_columns]

#                             values += [casc_id, permit_id, row_no]

#                             query = f"""
#                                 UPDATE {self.table}
#                                 SET {set_clause}
#                                 WHERE CASCId=%s
#                                   AND PermitId=%s
#                                   AND RowNo=%s
#                             """

#                             cursor.execute(query, values)
#                             updated_count += 1

#                     # ----------------------------
#                     # INSERT
#                     # ----------------------------
#                     else:

#                         columns = [
#                             col for col in self.allowed_columns
#                             if col in item
#                         ]

#                         values = [item[col] for col in columns]

#                         placeholders = ", ".join(["%s"] * len(columns))

#                         query = f"""
#                             INSERT INTO {self.table}
#                             ({', '.join(columns)})
#                             VALUES ({placeholders})
#                         """

#                         cursor.execute(query, values)
#                         inserted_count += 1

#                 connections['default'].commit()

#         except Exception as e:
#             return Response(
#                 {"error": f"Error saving CASC data: {str(e)}"},
#                 status=400
#             )

#         return Response(
#             {
#                 "message": f"{inserted_count} inserted, {updated_count} updated successfully"
#             },
#             status=201
#         )


class PostCascTable(APIView):
    table = "CommonCASCDtl"
    def post(self, request):
        payloads = request.data
        if not isinstance(payloads, list):
            payloads = [payloads]
        inserted_count = 0
        updated_count = 0
        try:
            with connections['default'].cursor() as cursor:
                for item in payloads:
                    item_no = item.get("ItemNo")
                    permit_id = item.get("PermitId")
                    row_no = item.get("RowNo")
                    casc_id = item.get("CASCId")
                    if not item_no or not permit_id or row_no is None:
                        continue
                    cursor.execute("""
                        SELECT COUNT(*)
                        FROM CommonCASCDtl
                        WHERE ItemNo=%s AND PermitId=%s AND RowNo=%s AND CASCId=%s
                    """, [item_no, permit_id, row_no, casc_id])
                    exists = cursor.fetchone()[0] > 0
                    if exists:
                        cursor.execute("""
                            UPDATE CommonCASCDtl
                            SET
                                ProductCode=%s,
                                Quantity=%s,
                                ProductUOM=%s,
                                CascCode1=%s,
                                CascCode2=%s,
                                CascCode3=%s,
                                TouchUser=%s,
                                TouchTime=%s,
                                EndUserDes=%s,
                                CASCId=%s
                            WHERE ItemNo=%s AND PermitId=%s AND RowNo=%s AND CASCId=%s
                        """, [
                            item.get("ProductCode"),
                            item.get("Quantity"),
                            item.get("ProductUOM"),
                            item.get("CascCode1"),
                            item.get("CascCode2"),
                            item.get("CascCode3"),
                            item.get("TouchUser"),
                            item.get("TouchTime"),
                            item.get("EndUserDes"),
                            casc_id,
                            item_no,
                            permit_id,
                            row_no,
                            casc_id
                        ])
                        updated_count += 1
                    else:
                        cursor.execute("""
                            INSERT INTO CommonCASCDtl
                            (
                                ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                PermitId, MessageType,
                                TouchUser, TouchTime, EndUserDes, CASCId
                            )
                            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                        """, [
                            item_no,
                            item.get("ProductCode"),
                            item.get("Quantity"),
                            item.get("ProductUOM"),
                            row_no,
                            item.get("CascCode1"),
                            item.get("CascCode2"),
                            item.get("CascCode3"),
                            permit_id,
                            item.get("MessageType"),
                            item.get("TouchUser"),
                            item.get("TouchTime"),
                            item.get("EndUserDes"),
                            casc_id
                        ])
                        inserted_count += 1
                connections['default'].commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({
            "inserted": inserted_count,
            "updated": updated_count
        }, status=200)

class EditCascByPermitId(APIView):
    table = "CommonCASCDtl"
    def put(self, request, permit_id):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","PermitId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(permit_id)
        query = f"UPDATE {self.table} SET {set_clause} WHERE PermitId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with PermitId {permit_id} updated successfully"})

#---------CPC-------

class GetCommonCpcTabel(APIView):
    table = "CommonCPCDtl"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonCpcByEditPermitId(APIView):
    table ="CommonCPCDtl"
    def get(self, request):
            try:
                permit_id = request.GET.get("PermitId")
                if not permit_id:
                    return Response(
                        {"error": "PermitId is required"},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
                rows = SqlDb.execute_query(query, [permit_id])
                if not rows:
                    # return Response(
                    #     {"error": "Record not found"},
                    #     status=status.HTTP_404_NOT_FOUND
                    # )        
                    return Response([], status=200)      
                return Response(rows, status=status.HTTP_200_OK)
            except Exception as e:
                return Response(
                    {"error": str(e)},
                    status=status.HTTP_500_INTERNAL_SERVER_ERROR
                )



class GetCommonCpcByPermitId(APIView):
    table = "CommonCPCDtl"
    def get(self, request, permit_id):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
        data = SqlDb.execute_query(query, [permit_id])
        if not data:
            return Response(
                {"message": f"No records found for PermitId {permit_id}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteCpc(APIView):
    table = "CommonCPCDtl"
    def delete(self, request, permit_id):
        select_query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
        existing = SqlDb.execute_query(select_query, [permit_id])
        if not existing:
            return Response({"error": f"No record found with PermitId {permit_id}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE PermitId = %s"
        try:
            SqlDb.execute_query(delete_query, [permit_id])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with PermitId {permit_id} deleted successfully"})



# class PostCpcTable(APIView):
#     table = "CommonCPCDtl"
#     allowed_columns = [
#         "PermitId", "MessageType", "RowNo", "CPCType", "ProcessingCode1",
#         "ProcessingCode2", "ProcessingCode3", "TouchUser", "TouchTime"
#     ]
#     def post(self, request):
#         payloads = request.data
#         if not payloads:
#             return Response({"error": "No data provided"}, status=400)       

#         if not isinstance(payloads, list):
#             payloads = [payloads]
#         permit_id = None
#         action = "processed"
#         try:
#             with connections['default'].cursor() as cursor:
#                 for item in payloads:
#                     if not isinstance(item, dict):
#                         item = dict(item)                  
#                     item.pop("Id", None) 
#                     columns = [k for k in item.keys() if k in self.allowed_columns]
#                     if not columns:
#                         continue

#                     permit_id = item.get("PermitId")
#                     row_no = item.get("RowNo")
#                     cpc_type = item.get("CPCType")

#                     if not permit_id or not row_no:
#                         return Response(
#                             {"error": "PermitId and RowNo are required for each CPC record"}, 
#                             status=400
#                         )

#                     # Check if this specific CPC row already exists for this permit
#                     cursor.execute(
#                         f"SELECT COUNT(*) FROM {self.table} WHERE PermitId=%s AND RowNo=%s AND CPCType=%s",
#                         [permit_id, row_no, cpc_type]
#                     )
#                     exists = cursor.fetchone()[0] > 0

#                     if exists:
#                         # UPDATE LOGIC
#                         update_columns = [col for col in columns if col not in ["PermitId", "RowNo", "CPCType"]]
#                         if update_columns:
#                             set_clause = ", ".join([f"{col}=%s" for col in update_columns])
#                             values = [item[col] for col in update_columns]
#                             values += [permit_id, row_no, cpc_type]
                            
#                             query = f"UPDATE {self.table} SET {set_clause} WHERE PermitId=%s AND RowNo=%s AND CPCType=%s"
#                             cursor.execute(query, values)
#                     else:
#                         # INSERT LOGIC
#                         placeholders = ", ".join(["%s"] * len(columns))
#                         values = [item[col] for col in columns]
#                         query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
#                         cursor.execute(query, values)

#                 connections['default'].commit()

#         except Exception as e:
#             return Response({"error": f"Database Error: {str(e)}"}, status=400)

#         return Response({
#             "Result": f"CPC data {action} successfully",
#             "PermitId": permit_id
#         }, status=201)

class PostCpcTable(APIView):
    table = "CommonCPCDtl"
    allowed_columns = [
        "PermitId", "MessageType", "RowNo", "CPCType", "ProcessingCode1",
        "ProcessingCode2", "ProcessingCode3", "TouchUser", "TouchTime"
    ]
    def post(self, request):
        payloads = request.data
        if not isinstance(payloads, list):
            payloads = [payloads] if payloads else []

        # Collect all distinct PermitIds present in the payload
        permit_ids = list({item.get("PermitId") for item in payloads if item.get("PermitId")})

        # If frontend explicitly requests a full replace for a permit
        # even with an empty payload, pass PermitId via query param.
        query_permit_id = request.GET.get("PermitId")
        if query_permit_id and query_permit_id not in permit_ids:
            permit_ids.append(query_permit_id)

        try:
            with connections['default'].cursor() as cursor:
                # ── STEP 1: Full replace — delete existing CPC rows first ──
                for pid in permit_ids:
                    cursor.execute(
                        f"DELETE FROM {self.table} WHERE PermitId=%s",
                        [pid]
                    )

                # ── STEP 2: Insert current (checked) CPC rows ──
                for item in payloads:
                    if not isinstance(item, dict):
                        item = dict(item)
                    item.pop("Id", None)
                    columns = [k for k in item.keys() if k in self.allowed_columns]
                    if not columns:
                        continue

                    permit_id = item.get("PermitId")
                    row_no = item.get("RowNo")
                    if not permit_id or not row_no:
                        return Response(
                            {"error": "PermitId and RowNo are required for each CPC record"},
                            status=400
                        )

                    placeholders = ", ".join(["%s"] * len(columns))
                    values = [item[col] for col in columns]
                    query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
                    cursor.execute(query, values)

                connections['default'].commit()

        except Exception as e:
            return Response({"error": f"Database Error: {str(e)}"}, status=400)

        return Response({
            "Result": "CPC data replaced successfully",
            "PermitIds": permit_ids
        }, status=201)


class EditCpcByPermitId(APIView):
    table = "CommonCPCDtl"
    def put(self, request, permit_id):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","PermitId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(permit_id)
        query = f"UPDATE {self.table} SET {set_clause} WHERE PermitId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with PermitId {permit_id} updated successfully"})

#----------Container--------

class GetCommonContainerTabel(APIView):
    table = "CommonContainerDtl"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetContainerByEditPermitId(APIView):
    table = "CommonContainerDtl"
    def get(self, request):
            try:
                permit_id = request.GET.get("PermitId")
                if not permit_id:
                    return Response(
                        {"error": "PermitId is required"},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
                rows = SqlDb.execute_query(query, [permit_id])
                if not rows:
                    # return Response(
                    #     {"error": "Record not found"},
                    #     status=status.HTTP_404_NOT_FOUND
                    # ) 
                    return Response([], status=200)              
                return Response(rows, status=status.HTTP_200_OK)
            except Exception as e:
                return Response(
                    {"error": str(e)},
                    status=status.HTTP_500_INTERNAL_SERVER_ERROR
                )

class GetCommonContainerByPermitId(APIView):
    table = "CommonContainerDtl"
    def get(self, request, permit_id):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
        data = SqlDb.execute_query(query, [permit_id])
        if not data:
            return Response(
                {"message": f"No records found for PermitId {permit_id}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)

class DeleteContainer(APIView):
    table = "CommonContainerDtl"
    def post(self, request):
        permit_id = request.data.get("PermitId")
        row_no = request.data.get("RowNo")
        if not permit_id or not row_no:
            return Response({"error": "PermitId and RowNo required"}, status=400)
        try:
            with connections['default'].cursor() as cursor:
                cursor.execute(
                    f"DELETE FROM {self.table} WHERE PermitId=%s AND RowNo=%s",
                    [permit_id, row_no]
                )
                cursor.execute(
                    f"""
                    UPDATE {self.table}
                    SET RowNo = RowNo - 1
                    WHERE PermitId=%s AND RowNo > %s
                    """,
                    [permit_id, row_no]
                )
                connections['default'].commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"Result": "Container deleted and rows reordered"})

class PostContainerTable(APIView):
    table = "CommonContainerDtl"
    allowed_columns = {
        "PermitId", "RowNo", "ContainerNo", "Size", "Weight",
        "SealNo", "MessageType", "TouchUser", "TouchTime"
    }

    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        # Remove Id column if exists
        item.pop("Id", None)

        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert/update"}, status=400)

        permit_id = item.get("PermitId")
        row_no = item.get("RowNo")
        if not permit_id or not row_no:
            return Response({"error": "PermitId and RowNo are required"}, status=400)

        try:
            with connections['default'].cursor() as cursor:
                # Check if this container already exists (update case)
                cursor.execute(
                    f"SELECT COUNT(*) FROM {self.table} WHERE PermitId=%s AND RowNo=%s",
                    [permit_id, row_no]
                )
                exists = cursor.fetchone()[0] > 0

                if exists:
                    # UPDATE existing row
                    update_columns = [col for col in columns if col not in ["PermitId", "RowNo"]]
                    set_clause = ", ".join([f"{col}=%s" for col in update_columns])
                    values = [item[col] for col in update_columns] + [permit_id, row_no]
                    update_query = f"UPDATE {self.table} SET {set_clause} WHERE PermitId=%s AND RowNo=%s"
                    cursor.execute(update_query, values)
                    action = "updated"
                else:
                    # INSERT new row
                    placeholders = ", ".join(["%s"] * len(columns))
                    values = [item[col] for col in columns]
                    insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
                    cursor.execute(insert_query, values)
                    action = "inserted"

                connections['default'].commit()

        except Exception as e:
            return Response({"error": f"Error saving record: {str(e)}"}, status=400)

        # Optionally, fetch all container records for this permit
        try:
            with connections['default'].cursor() as cursor:
                fetch_columns = ", ".join(self.allowed_columns)
                cursor.execute(f"SELECT {fetch_columns} FROM {self.table} WHERE PermitId=%s", [permit_id])
                rows = cursor.fetchall()
                records = [
                    dict(zip(self.allowed_columns, row))
                    for row in rows
                ]
        except Exception as e:
            records = []

        return Response({
            "Result": f"Container {action} successfully",
            "Records": records
        }, status=201)

class EditContainerByPermitId(APIView):
    table = "CommonContainerDtl"
    def put(self, request, permit_id):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","PermitId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(permit_id)
        query = f"UPDATE {self.table} SET {set_clause} WHERE PermitId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with PermitId {permit_id} updated successfully"})

# ------------Check Duplicate HBL/HAWB -------------------

class CheckDuplicateHawb(APIView):
    def get(self, request):
        try:
            hbl_value    = (request.GET.get("HBL") or "").strip()
            in_hawb_value  = (request.GET.get("INHAWB") or "").strip()
            out_hawb_value = (request.GET.get("outHAWB") or "").strip()
            permit_id = request.GET.get("PermitId", "")

            if not hbl_value and not in_hawb_value and not out_hawb_value:
                return Response({"duplicate": False})

            duplicates_found = {}

            def check_field(field_name, value):
                if not value:
                    return
                # Split comma-separated values the same way the frontend does
                value_list = [v.strip() for v in value.split(",") if v.strip()]
                for v in value_list:
                    dupe = SqlDb.execute_query(
                        f"""
                        SELECT PermitId FROM CommonHeaderTbl
                        WHERE {field_name} = %s AND PermitId != %s
                        """,
                        [v, permit_id]
                    )
                    if dupe:
                        duplicates_found.setdefault(field_name, []).append(v)

            check_field("HBL", hbl_value)
            check_field("INHAWB", in_hawb_value)
            check_field("outHAWB", out_hawb_value)

            if duplicates_found:
                messages = []
                if "HBL" in duplicates_found:
                    messages.append(f"Duplicate HBL/HAWB found: {', '.join(duplicates_found['HBL'])}")
                if "INHAWB" in duplicates_found:
                    messages.append(f"Duplicate In-HAWB found: {', '.join(duplicates_found['INHAWB'])}")
                if "outHAWB" in duplicates_found:
                    messages.append(f"Duplicate Out-HAWB found: {', '.join(duplicates_found['outHAWB'])}")

                return Response({
                    "duplicate": True,
                    "message": " | ".join(messages),
                    "fields": duplicates_found
                })

            return Response({"duplicate": False})

        except Exception as e:
            return Response({"error": str(e)}, status=500)

#-------------House Code------------------

class GetCommonHouseItemCode(APIView):
    table = "CommonhouseItemCode"
    def get(self, request):
        query = f"SELECT * FROM {self.table} ORDER BY HouseCode"
        data = SqlDb.execute_query(query)
        return Response(data)

class GetCommonHouseItemCodeByHouseCode(APIView):
    table = "CommonhouseItemCode"
    def get(self, request, housecode):
        query = f"SELECT * FROM {self.table} WHERE HouseCode = %s"
        data = SqlDb.execute_query(query, [housecode])
        if not data:
            return Response(
                {"message": f"No records found for HouseCode {housecode}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)

class DeleteCommonHouseItemCode(APIView):
    table = "CommonhouseItemCode"
    def delete(self, request, housecode):
        select_query = f"SELECT * FROM {self.table} WHERE HouseCode = %s"
        existing = SqlDb.execute_query(select_query, [housecode])
        if not existing:
            return Response(
                {"error": f"No record found with HouseCode {housecode}"},
                status=404
            )
        delete_query = f"DELETE FROM {self.table} WHERE HouseCode = %s"
        try:
            SqlDb.execute_query(delete_query, [housecode])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response(
            {"message": f"Record with HouseCode {housecode} deleted successfully"}
        )

class PostCommonHouseItemCode(APIView):
    table = "CommonhouseItemCode"
    allowed_columns = {
        "HouseCode","HSCode","Description","Brand","Model","DGIndicator","TouchUser","TouchTime","DeclType","ProductCode"}
    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        item.pop("Id", None)
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        housecode = item.get("HouseCode")
        if not housecode:
            return Response({"error": "HouseCode is required"}, status=400)
        select_query = f"SELECT HouseCode FROM {self.table} WHERE HouseCode = %s"
        existing = SqlDb.execute_query(select_query, [housecode])
        if existing:
            return Response(
                {"Result": "HouseCode Already Exists ...!"},
                status=400
            )
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"""
            INSERT INTO {self.table} ({', '.join(columns)})
            VALUES ({placeholders})
        """
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response(
                    {"error": f"Error inserting record: {str(e)}"},
                    status=400
                )
        return Response(
            {"Result": "Record Saved Successfully ...!"},
            status=201
        )


class EditCommonHouseItemCode(APIView):
    table = "CommonhouseItemCode"

    def put(self, request, housecode):

        payload = request.data

        if not payload:
            return Response(
                {"error": "No data provided"},
                status=status.HTTP_400_BAD_REQUEST
            )
        for key in ["Id", "HouseCode"]:
            payload.pop(key, None)

        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(housecode)

        query = f"""
            UPDATE {self.table}
            SET {set_clause}
            WHERE HouseCode = %s
        """

        try:
            SqlDb.execute_query(query, values)
        except Exception as e:
            return Response(
                {"error": str(e)},
                status=status.HTTP_400_BAD_REQUEST
            )

        return Response(
            {"message": f"Record with HouseCode {housecode} updated successfully"}
        )



#-------------importer--------------------
class GetCommonImporterTabel(APIView):
    table = "CommonImporter"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table} ORDER BY Code")
        return Response(data)

class GetCommonImporterByCode(APIView):
    table = "CommonImporter"
    def get(self, request, code):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE Code = %s"
        data = SqlDb.execute_query(query, [code])
        if not data:
            return Response(
                {"message": f"No records found for Code {code}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteImporter(APIView):
    table = "CommonImporter"
    def delete(self, request, code):
        select_query = f"SELECT * FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if not existing:
            return Response({"error": f"No record found with Code {code}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE Code = %s"
        try:
            SqlDb.execute_query(delete_query, [code])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with Code {code} deleted successfully"})


class PostImporterTable(APIView):
    table = "CommonImporter"
    allowed_columns = {"Code","Name","Name1","CRUEI","TouchUser","TouchTime","Status","MES","APS"}

    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        # Remove Id column
        item.pop("Id", None)
        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        # Check if Code exists
        code = item.get("Code")
        if not code:
            return Response({"error": "Code is required"}, status=400)
        select_query = f"SELECT Code FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "Code Already Exists ...!"}, status=400)
        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
        # INSERT: your execute_query is only for SELECT; need a separate commit
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        # Fetch all active records
        active_query = f"SELECT Code, Name, Name1, CRUEI FROM {self.table} WHERE Status = 'Active'"
        active_records = SqlDb.execute_query(active_query)
        return Response({
            "Result": "Record Saved Successfully ...!",
            # "Importer": active_records
        }, status=201)


class EditImporterByCode(APIView):
    table = "CommonImporter"
    def put(self, request, code):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","Code"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(code)
        query = f"UPDATE {self.table} SET {set_clause} WHERE Code = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with Code {code} updated successfully"})


#-------------Handling Agent--------------------
class GetCommonHandlingAgentTabel(APIView):
    table = "CommonHandingAgent"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table} ORDER BY Code")
        return Response(data)

class GetCommonHandlingAgentByCode(APIView):
    table = "CommonHandingAgent"
    def get(self, request, code):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE Code = %s"
        data = SqlDb.execute_query(query, [code])
        if not data:
            return Response(
                {"message": f"No records found for Code {code}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteHandlingAgent(APIView):
    table = "CommonHandingAgent"
    def delete(self, request, code):
        select_query = f"SELECT * FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if not existing:
            return Response({"error": f"No record found with Code {code}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE Code = %s"
        try:
            SqlDb.execute_query(delete_query, [code])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with Code {code} deleted successfully"})


class PostHandlingAgentTable(APIView):
    table = "CommonHandingAgent"
    allowed_columns = {"Code","Name","Name1","CRUEI","TouchUser","TouchTime","Status"}

    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        # Remove Id column
        item.pop("Id", None)
        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        # Check if Code exists
        code = item.get("Code")
        if not code:
            return Response({"error": "Code is required"}, status=400)
        select_query = f"SELECT Code FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "Code Already Exists ...!"}, status=400)
        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
        # INSERT: your execute_query is only for SELECT; need a separate commit
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        # Fetch all active records
        active_query = f"SELECT Code, Name, Name1, CRUEI FROM {self.table} WHERE Status = 'Active'"
        active_records = SqlDb.execute_query(active_query)
        return Response({
            "Result": "Record Saved Successfully ...!",
            # "Importer": active_records
        }, status=201)


class EditHandlingAgentByCode(APIView):
    table = "CommonHandingAgent"
    def put(self, request, code):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","Code"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(code)
        query = f"UPDATE {self.table} SET {set_clause} WHERE Code = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with Code {code} updated successfully"})



#---------------- Freight Forwarder ----------------#
class GetCommonFreightForwarderTable(APIView):
    table = "CommonFreightForwarder"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table} ORDER BY Code")
        return Response(data)


class GetFreightForwarderByCode(APIView):
    table = "CommonFreightForwarder"

    def get(self, request, code):
        """Return a record by Code"""
        query = f"SELECT * FROM {self.table} WHERE Code = %s"
        data = SqlDb.execute_query(query, [code])
        if not data:
            return Response(
                {"message": f"No records found for Code {code}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteFreightForwarder(APIView):
    table = "CommonFreightForwarder"

    def delete(self, request, code):
        """Delete a record by Code"""
        select_query = f"SELECT * FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if not existing:
            return Response(
                {"error": f"No record found with Code {code}"}, 
                status=404
            )
        delete_query = f"DELETE FROM {self.table} WHERE Code = %s"
        try:
            SqlDb.execute_query(delete_query, [code])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with Code {code} deleted successfully"})


class PostFreightForwarderTable(APIView):
    table = "CommonFreightForwarder"
    allowed_columns = {"Code","Name","Name1","CRUEI","TouchUser","TouchTime","Status","MES","APS"}
    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        item.pop("Id", None)
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        code = item.get("Code")
        if not code:
            return Response({"error": "Code is required"}, status=400)
        select_query = f"SELECT Code FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "Code Already Exists ...!"}, status=400)
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        active_query = f"SELECT Code, Name, Name1, CRUEI FROM {self.table} WHERE Status = 'Active'"
        active_records = SqlDb.execute_query(active_query)
        return Response({
            "Result": "Record Saved Successfully ...!",
        }, status=201)



class EditFreightForwarderByCode(APIView):
    table = "CommonFreightForwarder"

    def put(self, request, code):
        """Update a record by Code"""
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        
        # Remove non-updatable fields
        for key in ["Id","Code"]:
            payload.pop(key, None)
        
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(code)
        query = f"UPDATE {self.table} SET {set_clause} WHERE Code = %s"
        try:
           result = SqlDb.execute_query(query, values)   
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with Code {code} updated successfully"})

#CLAIMANT PARTY

class GetCommonClaimantPartyTable(APIView):
    table = "CommonClaimantParty"

    def get(self, request):
        """Return all records ordered by Id"""
        data = SqlDb.execute_query(f"SELECT * FROM {self.table} ORDER BY Id")
        return Response(data)


class GetClaimantPartyById(APIView):
    table = "CommonClaimantParty"

    def get(self, request, id):
        """Return a record by Id"""
        query = f"SELECT * FROM {self.table} WHERE Id = %s"
        data = SqlDb.execute_query(query, [id])
        if not data:
            return Response(
                {"message": f"No records found for Id {id}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteClaimantParty(APIView):
    table = "CommonClaimantParty"

    def delete(self, request, id):
        """Delete a record by Id"""
        select_query = f"SELECT * FROM {self.table} WHERE Id = %s"
        existing = SqlDb.execute_query(select_query, [id])
        if not existing:
            return Response({"error": f"No record found with Id {id}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE Id = %s"
        try:
            SqlDb.execute_query(delete_query, [id])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with Id {id} deleted successfully"})


class PostClaimantPartyTable(APIView):
    table = "CommonClaimantParty"
    allowed_columns = {"Name","Name1","CRUEI","ClaimantName","ClaimantName1","ClaimantCode","TouchUser","TouchTime","Name2","Status"}
    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        # Remove Id column
        item.pop("Id", None)
        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        # Check if Code exists
        code = item.get("ClaimantCode")
        if not code:
            return Response({"error": "ClaimantCode is required"}, status=400)
        select_query = f"SELECT ClaimantCode FROM {self.table} WHERE ClaimantCode = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "ClaimantCode Already Exists ...!"}, status=400)
        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
        # INSERT: your execute_query is only for SELECT; need a separate commit
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        # Fetch all active records
        active_query = f"SELECT Name, Name1, CRUEI,ClaimantName,ClaimantName1,ClaimantCode,Name2 FROM {self.table} WHERE Status = 'Active'"
        active_records = SqlDb.execute_query(active_query)
        return Response({
            "Result": "Record Saved Successfully ...!",
            # "ClaimantParty": active_records
        }, status=201)



class EditClaimantPartyById(APIView):
    table = "CommonClaimantParty"

    def put(self, request, id):
        """Update a record by Id"""
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)

        # Remove non-updatable fields
        for key in ["Id"]:
            payload.pop(key, None)

        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(id)

        query = f"UPDATE {self.table} SET {set_clause} WHERE Id = %s"
        try:
            SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with Id {id} updated successfully"})


#--------------Exporter---------------
class GetCommonExporterTabel(APIView):
    table = "CommonExporter"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonExporterByCode(APIView):
    table = "CommonExporter"
    def get(self, request, code):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE Code = %s"
        data = SqlDb.execute_query(query, [code])
        if not data:
            return Response(
                {"message": f"No records found for Code {code}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteExporter(APIView):
    table = "CommonExporter"
    def delete(self, request, code):
        select_query = f"SELECT * FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if not existing:
            return Response({"error": f"No record found with Code {code}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE Code = %s"
        try:
            SqlDb.execute_query(delete_query, [code])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with Code {code} deleted successfully"})




class PostExporterTable(APIView):
    table ="CommonExporter"
    allowed_columns = {"Code","Name","Name1","CRUEI","Address","Address1","City","SubCode","Sub","Postal","Country","TouchUser","TouchTime","Status"}
    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        # Remove Id column
        item.pop("Id", None)
        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        # Check if Code exists
        code = item.get("Code")
        if not code:
            return Response({"error": "Code is required"}, status=400)
        select_query = f"SELECT Code FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "Code Already Exists ...!"}, status=400)
        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
        # INSERT: your execute_query is only for SELECT; need a separate commit
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        # Fetch all active records
        active_query = f"SELECT Name,Name1,CRUEI,Address,Address1,City,SubCode,Sub,Postal,Country FROM {self.table} WHERE Status = 'Active'"
        active_records = SqlDb.execute_query(active_query)
        return Response({
            "Result": "Record Saved Successfully ...!",
           }, status=201)


class EditExporterByCode(APIView):
    table = "CommonExporter"
    def put(self, request, code):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","Code"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(code)
        query = f"UPDATE {self.table} SET {set_clause} WHERE Code = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with Code {code} updated successfully"})


#-------------Inward---------------

class CommonInwardCarrierAgent(APIView):
    table = "CommonInwardCarrierAgent"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonInwardCarrierAgentByCode(APIView):
    table = "CommonInwardCarrierAgent"
    def get(self, request, code):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE Code = %s"
        data = SqlDb.execute_query(query, [code])
        if not data:
            return Response(
                {"message": f"No records found for Code {code}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteInwardCarrierAgent(APIView):
    table = "CommonInwardCarrierAgent"
    def delete(self, request, code):
        select_query = f"SELECT * FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if not existing:
            return Response({"error": f"No record found with Code {code}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE Code = %s"
        try:
            SqlDb.execute_query(delete_query, [code])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with Code {code} deleted successfully"})


class PostInwardCarrierAgentTable(APIView):
    table ="CommonInwardCarrierAgent"
    allowed_columns = {"Code","Name","Name1","CRUEI","TouchUser","TouchTime","Status"}
    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        # Remove Id column
        item.pop("Id", None)
        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        # Check if Code exists
        code = item.get("Code")
        if not code:
            return Response({"error": "Code is required"}, status=400)
        select_query = f"SELECT Code FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "Code Already Exists ...!"}, status=400)
        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
        # INSERT: your execute_query is only for SELECT; need a separate commit
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        # Fetch all active records
        active_query = f"SELECT Code, Name, Name1, CRUEI FROM {self.table} WHERE Status = 'Active'"
        active_records = SqlDb.execute_query(active_query)
        return Response({
            "Result": "Record Saved Successfully ...!",
            # "Inward": active_records
        }, status=201)


class EditInwardCarrierAgentByCode(APIView):
    table = "CommonInwardCarrierAgent"
    def put(self, request, code):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","Code"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(code)
        query = f"UPDATE {self.table} SET {set_clause} WHERE Code = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with Code {code} updated successfully"})

#------------Outward----------------

class CommonOutwardCarrierAgent(APIView):
    table = "CommonOutwardCarrierAgent"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonOutwardCarrierAgentByCode(APIView):
    table = "CommonOutwardCarrierAgent"
    def get(self, request, code):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE Code = %s"
        data = SqlDb.execute_query(query, [code])
        if not data:
            return Response(
                {"message": f"No records found for Code {code}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteOutwardCarrierAgent(APIView):
    table = "CommonOutwardCarrierAgent"
    def delete(self, request, code):
        select_query = f"SELECT * FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if not existing:
            return Response({"error": f"No record found with Code {code}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE Code = %s"
        try:
            SqlDb.execute_query(delete_query, [code])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with Code {code} deleted successfully"})




class PostOutwardCarrierAgentTable(APIView):
    table ="CommonOutwardCarrierAgent"
    allowed_columns = {"Code","Name","Name1","CRUEI","TouchUser","TouchTime","Status"}
    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)
        # Remove Id column
        item.pop("Id", None)
        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)
        # Check if Code exists
        code = item.get("Code")
        if not code:
            return Response({"error": "Code is required"}, status=400)
        select_query = f"SELECT Code FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "Code Already Exists ...!"}, status=400)
        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"
        # INSERT: your execute_query is only for SELECT; need a separate commit
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        # Fetch all active records
        active_query = f"SELECT Name,Name1,CRUEI FROM {self.table} WHERE Status = 'Active'"
        active_records = SqlDb.execute_query(active_query)
        return Response({
            "Result": "Record Saved Successfully ...!",
           }, status=201)


class EditOutwardCarrierAgentByCode(APIView):
    table = "CommonOutwardCarrierAgent"
    def put(self, request, code):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","Code"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(code)
        query = f"UPDATE {self.table} SET {set_clause} WHERE Code = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with Code {code} updated successfully"})


#--------------Consignee-----------------

class CommonConsigneeTable(APIView):
    table = "CommonConsignee"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonConsigneeAgentByCode(APIView):
    table = "CommonConsignee"
    def get(self, request, consigneecode):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE ConsigneeCode = %s"
        data = SqlDb.execute_query(query, [consigneecode])
        if not data:
            return Response(
                {"message": f"No records found for ConsigneeCode {consigneecode}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteConsigneeAgentByCode(APIView):
    table = "CommonConsignee"
    def delete(self, request, consigneecode):
        select_query = f"SELECT * FROM {self.table} WHERE ConsigneeCode = %s"
        existing = SqlDb.execute_query(select_query, [consigneecode])
        if not existing:
            return Response({"error": f"No record found with ConsigneeCode {consigneecode}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE ConsigneeCode = %s"
        try:
            SqlDb.execute_query(delete_query, [consigneecode])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with ConsigneeCode {consigneecode} deleted successfully"})




class PostCongineeTable(APIView):
    table ="CommonConsignee"
    allowed_columns={ "ConsigneeCode","ConsigneeName","ConsigneeName1","ConsigneeCRUEI","ConsigneeAddress","ConsigneeAddress1","ConsigneeCity","ConsigneeSub","ConsigneeSubDivi","ConsigneePostal","TouchUser","ConsigneeCountry","TouchTime","Status"}

    def post(self, request):
        item = request.data

        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        # Remove Id column
        item.pop("Id", None)

        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]

        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)

        # Check if ConsigneeCode exists
        code = item.get("ConsigneeCode")

        if not code:
            return Response({"error": "ConsigneeCode is required"}, status=400)

        select_query = f"""
            SELECT ConsigneeCode
            FROM {self.table}
            WHERE ConsigneeCode = %s
        """

        existing = SqlDb.execute_query(select_query, [code])

        if existing:
            return Response(
                {"Result": "ConsigneeCode Already Exists ...!"},
                status=400
            )

        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]

        insert_query = f"""
            INSERT INTO {self.table}
            ({', '.join(columns)})
            VALUES ({placeholders})
        """

        # Insert record
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()

            except Exception as e:
                return Response(
                    {"error": f"Error inserting record: {str(e)}"},
                    status=400
                )

        # Fetch active records
        active_query = f"""
            SELECT
                ConsigneeCode,ConsigneeName,ConsigneeName1,ConsigneeCRUEI,ConsigneeAddress,ConsigneeAddress1,ConsigneeCity,ConsigneeSub,ConsigneeSubDivi,ConsigneePostal,ConsigneeCountry
            FROM {self.table}
            WHERE Status = 'Active'
        """

        active_records = SqlDb.execute_query(active_query)

        return Response({
            "Result": "Record Saved Successfully ...!",
        }, status=201)


class EditCongineeCode(APIView):
    table = "CommonConsignee"
    def put(self, request, consigneecode):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","ConsigneeCode"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(consigneecode)
        query = f"UPDATE {self.table} SET {set_clause} WHERE ConsigneeCode = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with ConsigneeCode {consigneecode} updated successfully"})


        
#--------------EndUser-----------------

class CommonEndUserTable(APIView):
    table = "CommonEndUser"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonEndUserByCode(APIView):
    table = "CommonEndUser"
    def get(self, request, EndUserCode):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE EndUserCode = %s"
        data = SqlDb.execute_query(query, [EndUserCode])
        if not data:
            return Response(
                {"message": f"No records found for EndUserCode {EndUserCode}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteEndUserByCode(APIView):
    table = "CommonEndUser"
    def delete(self, request, EndUserCode):
        select_query = f"SELECT * FROM {self.table} WHERE EndUserCode = %s"
        existing = SqlDb.execute_query(select_query, [consigneecode])
        if not existing:
            return Response({"error": f"No record found with EndUserCode {EndUserCode}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE EndUserCode = %s"
        try:
            SqlDb.execute_query(delete_query, [EndUserCode])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with EndUserCode {EndUserCode} deleted successfully"})



class PostEndUserTable(APIView):
    table = "CommonEndUser"
    allowed_columns = {
        "EndUserCode", "EndUserName", "EndUserName1", "EndUserCRUEI",
        "EndUserAddress", "EndUserAddress1", "EndUserCity",
        "EndUserSubCode", "EndUserSub",
        "EndUserPostal", "EndUserCountry",
        "TouchUser", "TouchTime", "Status"
    }

    def post(self, request):
        item = request.data

        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        item.pop("Id", None)

        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)

        code = item.get("EndUserCode")
        if not code:
            return Response({"error": "EndUserCode is required"}, status=400)

        select_query = f"SELECT EndUserCode FROM {self.table} WHERE EndUserCode = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "EndUserCode Already Exists ...!"}, status=400)

        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"

        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)

        return Response({"Result": "Record Saved Successfully ...!"}, status=201)
class EditEndUserCode(APIView):
    table = "CommonEndUser"
    def put(self, request, consigneecode):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","CommonEndUser"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(consigneecode)
        query = f"UPDATE {self.table} SET {set_clause} WHERE CommonEndUser = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with CommonEndUser {consigneecode} updated successfully"})

# ==========================Manufacturer=======================

class CommonManufacturerTable(APIView):
    table = "CommonManufacturer"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)


class GetCommonManufacturerByCode(APIView):
    table = "CommonManufacturer"
    def get(self, request, ManufacturerCode):
        query = f"SELECT * FROM {self.table} WHERE ManufacturerCode = %s"
        data = SqlDb.execute_query(query, [ManufacturerCode])
        if not data:
            return Response(
                {"message": f"No records found for ManufacturerCode {ManufacturerCode}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteManufacturerByCode(APIView):
    table = "CommonManufacturer"
    def delete(self, request, ManufacturerCode):
        select_query = f"SELECT * FROM {self.table} WHERE ManufacturerCode = %s"
        existing = SqlDb.execute_query(select_query, [ManufacturerCode])
        if not existing:
            return Response({"error": f"No record found with ManufacturerCode {ManufacturerCode}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE ManufacturerCode = %s"
        try:
            SqlDb.execute_query(delete_query, [ManufacturerCode])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with ManufacturerCode {ManufacturerCode} deleted successfully"})


class PostManufacturerTable(APIView):
    table = "CommonManufacturer"
    allowed_columns = {
        "ManufacturerCode", "ManufacturerName", "ManufacturerName1",
        "ManufacturerCRUEI", "ManufacturerAddress", "ManufacturerAddress1",
        "ManufacturerCity", "ManufacturerSubDivi", "ManufacturerSub",
        "ManufacturerPostal", "ManufacturerCountry",
        "TouchUser", "TouchTime", "Status"
    }

    def post(self, request):
        item = request.data

        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        # Remove Id column
        item.pop("Id", None)

        # Filter allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]

        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)

        # Check if ManufacturerCode exists
        code = item.get("ManufacturerCode")

        if not code:
            return Response({"error": "ManufacturerCode is required"}, status=400)

        select_query = f"""
            SELECT ManufacturerCode
            FROM {self.table}
            WHERE ManufacturerCode = %s
        """

        existing = SqlDb.execute_query(select_query, [code])

        if existing:
            return Response(
                {"Result": "ManufacturerCode Already Exists ...!"},
                status=400
            )

        # Prepare insert query
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]

        insert_query = f"""
            INSERT INTO {self.table}
            ({', '.join(columns)})
            VALUES ({placeholders})
        """

        # Insert record
        with connections[SqlDb.database_name].cursor() as cursor:
            try:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
            except Exception as e:
                return Response(
                    {"error": f"Error inserting record: {str(e)}"},
                    status=400
                )

        # Fetch active records
        active_query = f"""
            SELECT
                ManufacturerCode, ManufacturerName, ManufacturerName1,
                ManufacturerCRUEI, ManufacturerAddress, ManufacturerAddress1,
                ManufacturerCity, ManufacturerSubDivi, ManufacturerSub,
                ManufacturerPostal, ManufacturerCountry
            FROM {self.table}
            WHERE Status = 'Active'
        """

        active_records = SqlDb.execute_query(active_query)

        return Response({
            "Result": "Record Saved Successfully ...!",
        }, status=201)


class EditManufacturerCode(APIView):
    table = "CommonManufacturer"
    def put(self, request, ManufacturerCode):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id", "CommonManufacturer"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(ManufacturerCode)
        query = f"UPDATE {self.table} SET {set_clause} WHERE ManufacturerCode = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with ManufacturerCode {ManufacturerCode} updated successfully"})



# ================== CommonSupplierManufacturerPart ==================
# ================== Get all records ==================
class GetCommonSupplierManufacturerPart(APIView):
    table = "CommonSupplierManufacturerPart"

    def get(self, request):
        query = f"SELECT * FROM {self.table} ORDER BY Code"
        data = SqlDb.execute_query(query)
        return Response(data)


# ================== Get single record by Code ==================
class GetCommonSupplierManufacturerPartByCode(APIView):
    table = "CommonSupplierManufacturerPart"

    def get(self, request, code):
        query = f"SELECT * FROM {self.table} WHERE Code = %s"
        data = SqlDb.execute_query(query, [code])
        if not data:
            return Response(
                {"message": f"No records found for Code {code}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


# ================== Delete record ==================
class DeleteCommonSupplierManufacturerPart(APIView):
    table = "CommonSupplierManufacturerPart"

    def delete(self, request, code):
        select_query = f"SELECT * FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if not existing:
            return Response({"error": f"No record found with Code {code}"}, status=404)

        delete_query = f"DELETE FROM {self.table} WHERE Code = %s"
        try:
            with connections[SqlDb.database_name].cursor() as cursor:
                cursor.execute(delete_query, [code])
                connections[SqlDb.database_name].commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)

        return Response({"message": f"Record with Code {code} deleted successfully"})


# ================== Insert record ==================
class PostCommonSupplierManufacturerPart(APIView):
    table = "CommonSupplierManufacturerPart"
    allowed_columns = {"Code","Name","Name1","CRUEI","TouchUser","TouchTime","Status"}

    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        # Remove Id column if present
        item.pop("Id", None)

        # Filter only allowed columns
        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns to insert"}, status=400)

        # Check for Code
        code = item.get("Code")
        if not code:
            return Response({"error": "Code is required"}, status=400)

        select_query = f"SELECT Code FROM {self.table} WHERE Code = %s"
        existing = SqlDb.execute_query(select_query, [code])
        if existing:
            return Response({"Result": "Code Already Exists ...!"}, status=400)

        # Prepare insert
        placeholders = ", ".join(["%s"] * len(columns))
        values = [item[k] for k in columns]
        insert_query = f"INSERT INTO {self.table} ({', '.join(columns)}) VALUES ({placeholders})"

        try:
            with connections[SqlDb.database_name].cursor() as cursor:
                cursor.execute(insert_query, values)
                connections[SqlDb.database_name].commit()
        except Exception as e:
            return Response({"error": f"Error inserting record: {str(e)}"}, status=400)

        return Response({"Result": "Record Saved Successfully ...!"}, status=201)


# ================== Update record by Code ==================
class EditCommonSupplierManufacturerPartByCode(APIView):
    table = "CommonSupplierManufacturerPart"

    def put(self, request, code):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)

        # Remove Id and Code to prevent updating them
        for key in ["Id","Code"]:
            payload.pop(key, None)

        if not payload:
            return Response({"error": "No updatable fields provided"}, status=400)

        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(code)
        query = f"UPDATE {self.table} SET {set_clause} WHERE Code = %s"

        try:
            with connections[SqlDb.database_name].cursor() as cursor:
                cursor.execute(query, values)
                connections[SqlDb.database_name].commit()
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)

        return Response({"message": f"Record with Code {code} updated successfully"})
#--------------file-----------------


class GetFileTabel(APIView):
    table = "CommonFile"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonFileByEditPermitId(APIView):
    table ="CommonFile"
    def get(self, request):
        try:
            permit_id = request.GET.get("PermitId")
            if not permit_id:
                return Response(
                    {"error": "PermitId is required"},
                    status=status.HTTP_400_BAD_REQUEST
                )
            query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
            rows = SqlDb.execute_query(query, [permit_id])
            if not rows:
                return Response(
                    {"error": "Record not found"},
                    status=status.HTTP_404_NOT_FOUND
                )              
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as e:
            return Response(
                {"error": str(e)},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )


class GetCommonFileByPermitId(APIView):
    table = "CommonFile"
    def get(self, request, permit_id):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
        data = SqlDb.execute_query(query, [permit_id])
        if not data:
            return Response(
                {"message": f"No records found for PermitId {permit_id}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteFile(APIView):
    table = "CommonFile"
    def delete(self, request, permit_id, sno):
        try:
            print("PERMIT:", permit_id, "SNO:", sno)
            select_query = f"""
                SELECT * FROM {self.table}
                WHERE PermitId = %s AND Sno = %s
            """
            existing = SqlDb.execute_query(select_query, [permit_id, sno])
            print("EXISTING:", existing)
            if not existing:
                return Response({"error": "Record not found"}, status=404)
            row = existing[0]
            file_path = None
            for key in row.keys():
                if key.lower() == "filepath":
                    file_path = row[key]
                    break
            if file_path:
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                except Exception as e:
                    print("FILE DELETE ERROR:", e)
            delete_query = f"""
                DELETE FROM {self.table}
                WHERE PermitId = %s AND Sno = %s
            """
            SqlDb.execute_query(delete_query, [permit_id, sno])
            update_query = f"""
                UPDATE {self.table}
                SET Sno = Sno - 1
                WHERE PermitId = %s AND Sno > %s
            """
            SqlDb.execute_query(update_query, [permit_id, sno])
            SqlDb.commit()
            fetch_query = f"""
                SELECT Sno, Name, DocumentType, Size, filePath
                FROM {self.table}
                WHERE PermitId = %s
                ORDER BY Sno
            """
            records = SqlDb.execute_query(fetch_query, [permit_id])
            return Response({
                "message": "Deleted successfully",
                "Records": records
            })
        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": str(e)}, status=400)


class PostFileTable(APIView):
    table = "CommonFile"
    allowed_columns = [
        "Sno","Name","ContentType","Data","DocumentType",
        "PaymentId","TouchUser","TouchTime","filePath",
        "Size","PermitId","Type"
    ]
    def post(self, request):
        payload = request.data
        permit_id = payload.get("PermitId")
        sno = payload.get("Sno")
        if not permit_id or not sno:
            return Response({"error": "PermitId and Sno required"}, status=400)
        try:
            # ================= FILE SAVE =================
            file = request.FILES.get("file")
            file_path = ""
            if file:
                upload_dir = "D:\\KttProject\\kttproject\\FilePath"
                if not os.path.exists(upload_dir):
                    os.makedirs(upload_dir)
                file_path = os.path.join(upload_dir, file.name)
                with open(file_path, "wb+") as f:
                    for chunk in file.chunks():
                        f.write(chunk)
            with connections['default'].cursor() as cursor:
                # CHECK EXIST
                cursor.execute(
                    f"""SELECT COUNT(*) FROM {self.table} 
                        WHERE PermitId=%s AND Sno=%s""",
                    [permit_id, sno]
                )
                exists = cursor.fetchone()[0] > 0
                columns = []
                values = []
                for key in payload:
                    if key in self.allowed_columns:
                        columns.append(key)
                        values.append(payload.get(key))
                # add filepath
                if file_path:
                    columns.append("filePath")
                    values.append(file_path)
                # ================= UPDATE =================
                if exists:
                    update_cols = [
                        col for col in columns
                        if col not in ["PermitId", "Sno"]
                    ]
                    if update_cols:
                        set_clause = ", ".join([f"{c}=%s" for c in update_cols])
                        update_values = [payload.get(c) for c in update_cols]
                        update_values += [permit_id, sno]
                        cursor.execute(
                            f"""UPDATE {self.table}
                                SET {set_clause}
                                WHERE PermitId=%s AND Sno=%s""",
                            update_values
                        )
                    action = "updated"
                # ================= INSERT =================
                else:
                    placeholders = ", ".join(["%s"] * len(columns))
                    cursor.execute(
                        f"""INSERT INTO {self.table}
                            ({", ".join(columns)})
                            VALUES ({placeholders})""",
                        values
                    )
                    action = "inserted"
                connections['default'].commit()
                # ================= FETCH =================
                cursor.execute(
                    f"""SELECT {", ".join(self.allowed_columns)}
                        FROM {self.table}
                        WHERE PermitId=%s
                        ORDER BY Sno""",
                    [permit_id]
                )
                rows = cursor.fetchall()
                records = [
                    dict(zip(self.allowed_columns, row))
                    for row in rows
                ]
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({
            "Result": f"File {action} successfully",
            "Records": records
        }, status=201
        )

class EditFileByPermitId(APIView):
    table = "CommonFile"
    def put(self, request, permit_id):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","PermitId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(permit_id)
        query = f"UPDATE {self.table} SET {set_clause} WHERE PermitId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with PermitId {permit_id} updated successfully"})


#-----------PMT---------------
class GetPMTTabel(APIView):
    table = "CommonPMT"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonPMTByPermitNo(APIView):
    table = "CommonPMT"
    def get(self, request, permit_number):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE PermitNumber = %s"
        data = SqlDb.execute_query(query, [permit_number])
        if not data:
            return Response(
                {"message": f"No records found for PermitNumber {permit_number}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeletePmt(APIView):
    table = "CommonPMT"
    def delete(self, request, permit_number):
        select_query = f"SELECT * FROM {self.table} WHERE PermitNumber = %s"
        existing = SqlDb.execute_query(select_query, [permit_number])
        if not existing:
            return Response({"error": f"No record found with PermitNumber {permit_id}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE PermitNumber = %s"
        try:
            SqlDb.execute_query(delete_query, [permit_number])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with PermitNumber {permit_number} deleted successfully"})


class PostPmtTable(APIView):
    table ="CommonPMT"
    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)
        if not isinstance(payloads, list):
            payloads = [payloads]  # handle single-object case
        inserted_count = 0
        for item in payloads:
            if not isinstance(item, dict):
                item = dict(item)  # convert QueryDict / other mapping to dict
            item.pop("Id", None)  # safe removal of Id
            columns = ", ".join(item.keys())
            placeholders = ", ".join(["%s"] * len(item))
            values = list(item.values())
            query = f"""INSERT INTO {self.table} ({columns}) VALUES ({placeholders})"""
            #print("query:",query)
            try:
                SqlDb.execute_query(query, values)
                inserted_count += 1
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        SqlDb.commit()
        return Response(
            {"message": f"{inserted_count} record(s) inserted successfully"},
            status=201
        )

class EditPmtByPermitNo(APIView):
    table = "CommonPMT"
    def put(self, request, permit_number):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","PermitNumber"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(permit_number)
        query = f"UPDATE {self.table} SET {set_clause} WHERE PermitNumber = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with PermitNumber {permit_number} updated successfully"})



#-----------------AmdPmt--------------------
class GetAMDPMTTabel(APIView):
    table = "CommonAMDPMT"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonAMDPMTByPermitNo(APIView):
    table = "CommonAMDPMT"
    def get(self, request, permit_number):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE PermitNumber = %s"
        data = SqlDb.execute_query(query, [permit_number])
        if not data:
            return Response(
                {"message": f"No records found for PermitNumber {permit_number}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteAMDPmt(APIView):
    table = "CommonAMDPMT"
    def delete(self, request, permit_number):
        select_query = f"SELECT * FROM {self.table} WHERE PermitNumber = %s"
        existing = SqlDb.execute_query(select_query, [permit_number])
        if not existing:
            return Response({"error": f"No record found with PermitNumber {permit_id}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE PermitNumber = %s"
        try:
            SqlDb.execute_query(delete_query, [permit_number])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with PermitNumber {permit_number} deleted successfully"})


class PostAMDPmtTable(APIView):
    table ="CommonAMDPMT"
    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)
        if not isinstance(payloads, list):
            payloads = [payloads]  # handle single-object case
        inserted_count = 0
        for item in payloads:
            if not isinstance(item, dict):
                item = dict(item)  # convert QueryDict / other mapping to dict
            item.pop("Id", None)  # safe removal of Id
            columns = ", ".join(item.keys())
            placeholders = ", ".join(["%s"] * len(item))
            values = list(item.values())
            query = f"""INSERT INTO {self.table} ({columns}) VALUES ({placeholders})"""
            #print("query:",query)
            try:
                SqlDb.execute_query(query, values)
                inserted_count += 1
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        SqlDb.commit()
        return Response(
            {"message": f"{inserted_count} record(s) inserted successfully"},
            status=201
        )

class EditAMDPmtByPermitNo(APIView):
    table = "CommonAMDPMT"
    def put(self, request, permit_number):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","PermitNumber"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(permit_number)
        query = f"UPDATE {self.table} SET {set_clause} WHERE PermitNumber = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with PermitNumber {permit_number} updated successfully"})


#-------------------RejectStatus----------------------

class CommonRejectStausTable(APIView):
    table = "CommonRejectStatus"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonRejectStautsByMsgId(APIView):
    table = "CommonRejectStatus"
    def get(self, request, msgId):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE MsgId = %s"
        data = SqlDb.execute_query(query, [msgId])
        if not data:
            return Response(
                {"message": f"No records found for MsgId {msgId}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteRejectStatusByMsgId(APIView):
    table = "CommonRejectStatus"
    def delete(self, request, msgId):
        select_query = f"SELECT * FROM {self.table} WHERE MsgId = %s"
        existing = SqlDb.execute_query(select_query, [msgId])
        if not existing:
            return Response({"error": f"No record found with MsgId {msgId}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE MsgId = %s"
        try:
            SqlDb.execute_query(delete_query, [msgId])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with MsgId {msgId} deleted successfully"})




class PostRejectStautsTable(APIView):
    table ="CommonRejectStatus"
    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)
        if not isinstance(payloads, list):
            payloads = [payloads]  # handle single-object case
        inserted_count = 0
        for item in payloads:
            if not isinstance(item, dict):
                item = dict(item)  # convert QueryDict / other mapping to dict
            item.pop("Id", None)  # safe removal of Id
            columns = ", ".join(item.keys())
            placeholders = ", ".join(["%s"] * len(item))
            values = list(item.values())
            query = f"""INSERT INTO {self.table} ({columns}) VALUES ({placeholders})"""
            #print("query:",query)
            try:
                SqlDb.execute_query(query, values)
                inserted_count += 1
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        SqlDb.commit()
        return Response(
            {"message": f"{inserted_count} record(s) inserted successfully"},
            status=201
        )


class EditRejectStatusByMsgId(APIView):
    table = "CommonRejectStatus"
    def put(self, request, msgId):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","MsgId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(msgId)
        query = f"UPDATE {self.table} SET {set_clause} WHERE MsgId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with MsgId {msgId} updated successfully"})


#------------------------ErrorStatus----------------------#

class CommonErrorStatusTable(APIView):
    table = "CommonErrorStatus"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonErrorStautsByMsgId(APIView):
    table = "CommonErrorStatus"
    def get(self, request, msgId):
        # Parameterized query to prevent SQL injection
        query = f"SELECT * FROM {self.table} WHERE MsgId = %s"
        data = SqlDb.execute_query(query, [msgId])
        if not data:
            return Response(
                {"message": f"No records found for MsgId {msgId}"},
                status=status.HTTP_404_NOT_FOUND
            )
        return Response(data)


class DeleteErrorStatusByMsgId(APIView):
    table = "CommonErrorStatus"
    def delete(self, request, msgId):
        select_query = f"SELECT * FROM {self.table} WHERE MsgId = %s"
        existing = SqlDb.execute_query(select_query, [msgId])
        if not existing:
            return Response({"error": f"No record found with MsgId {msgId}"}, status=404)
        delete_query = f"DELETE FROM {self.table} WHERE MsgId = %s"
        try:
            SqlDb.execute_query(delete_query, [msgId])
            SqlDb.commit()
        except Exception as e:
            return Response({"error": str(e)}, status=400)
        return Response({"message": f"Record with MsgId {msgId} deleted successfully"})


class PostErrorStautsTable(APIView):
    table ="CommonErrorStatus"
    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)
        if not isinstance(payloads, list):
            payloads = [payloads]  # handle single-object case
        inserted_count = 0
        for item in payloads:
            if not isinstance(item, dict):
                item = dict(item)  # convert QueryDict / other mapping to dict
            item.pop("Id", None)  # safe removal of Id
            columns = ", ".join(item.keys())
            placeholders = ", ".join(["%s"] * len(item))
            values = list(item.values())
            query = f"""INSERT INTO {self.table} ({columns}) VALUES ({placeholders})"""
            #print("query:",query)
            try:
                SqlDb.execute_query(query, values)
                inserted_count += 1
            except Exception as e:
                return Response({"error": f"Error inserting record: {str(e)}"}, status=400)
        SqlDb.commit()
        return Response(
            {"message": f"{inserted_count} record(s) inserted successfully"},
            status=201
        )


class EditErrorStatusByMsgId(APIView):
    table = "CommonErrorStatus"
    def put(self, request, msgId):
        payload = request.data
        if not payload:
            return Response({"error": "No data provided"}, status=status.HTTP_400_BAD_REQUEST)
        for key in ["Id","MsgId"]:
            payload.pop(key, None)
        set_clause = ", ".join([f"{col} = %s" for col in payload.keys()])
        values = list(payload.values())
        values.append(msgId)
        query = f"UPDATE {self.table} SET {set_clause} WHERE MsgId = %s"
        try:
            result = SqlDb.execute_query(query, values)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_400_BAD_REQUEST)
        return Response({"message": f"Record with MsgId {msgId} updated successfully"})
#-------------------------------------------------------------------------------------#
#07-02-26
#Commom Drop Down Tables for Api

#Commonmaster

class GetCommonMasterTable(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

#Currency
class GetCommonCurrenyTable(APIView):
    table = "Currency"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table} ORDER BY Currency")
        for item in data:
            if "CurrencyRate" in item and item["CurrencyRate"] is not None:
                item["CurrencyRate"] = format(Decimal(str(item["CurrencyRate"])), 'f')
        return Response(data)

#Country
class GetCommonCountryTable(APIView):
    table = "Country"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table} ORDER BY CountryCode")
        return Response(data)

#Declreant Company
class GetCommonDeclarantCompanyTable(APIView):
    table = "DeclarantCompany"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)
#Hscode
class GetCommonHsCodeTable(APIView):
    table = "Hscode"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)
# Hscode
class GetCommonHsCodeAndDescriptionTable(APIView):
    table = "Hscode"
    def get(self, request):
        query = f"""
            SELECT HSCode, Description
            FROM {self.table}
            ORDER BY HSCode
        """
        data = SqlDb.execute_query(query)
        return Response(data)

#ItemUOM
class GetUOMFromCommonHscode(APIView):
    table = "Hscode"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT UOM from {self.table}")
        return Response(data)


#Termtype
class GetTermTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 7 AND StatusID=1 ORDER BY Name")
        return Response(data)


#TotalOuterPack
class GetTotalOuterPackFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 10 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DECLARINGFOR
class GetDeclaringForFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 80 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DECLARINGFOR
class GetDeclaringForFromCommonMasterByOutandTranshipment(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 81 AND StatusID=1 ORDER BY Name")
        return Response(data)

#InwardTransportMode
class GetInwardTransportModeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 3 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DeclarationType
class GetDeclarationTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 13 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DeclarationType--inpayment
class GetDeclarationTypeFromCommonMasterForInpayment(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 1 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DeclarationType--innonpayment
class GetDeclarationTypeFromCommonMasterForInnonpayment(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 13 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DeclarationType--out
class GetDeclarationTypeFromCommonMasterForOut(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 15 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DeclartionType--Transhipment
class GetDeclarationTypeFromCommonMasterForTranshipment(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 18 AND StatusID=1 ORDER BY Name")
        return Response(data)


#DeclarationType--out
class GetCoTypeFromCommonMasterForOut(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 16 AND StatusID=1 ORDER BY Name")
        return Response(data)


#BgIndicator
class GetBgIndicatorFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 4 AND StatusID=1 ORDER BY Name")
        return Response(data)

#Preferntial
class GetPreferntialFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 11 AND StatusID=1 ORDER BY Name")
        return Response(data)

#VehicalType
class GetVehicalTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 20 AND StatusID=1 ORDER BY Name")
        return Response(data)

#EngineCapacity
class GetEngineCapacityFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 21 AND StatusID=1 ORDER BY Name")
        return Response(data)

#MakingLot
class GetMakingLotFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 12 AND StatusID=1 ORDER BY Name")
        return Response(data)

#CancelType
class GetCancelTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name,Description from {self.table} WHERE TypeId= 75 AND StatusID=1 ORDER BY Name")
        return Response(data)

#RefundType
class GetRefundTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name,Description from {self.table} WHERE TypeId= 77 AND StatusID=1 ORDER BY Name")
        return Response(data)

#RefundType
class GetReasonForRefundFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name,Description from {self.table} WHERE TypeId= 76 AND StatusID=1 ORDER BY Name")
        return Response(data)

#CoType
class GetCoTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 16 AND StatusID=1 ORDER BY Name")
        return Response(data)

#DocumentAttachmentType
class GetDocumentAttachTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 5 AND StatusID=1 ORDER BY Name")
        return Response(data)

#CertificateType
class GetCertificateTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 17 AND StatusID=1 ORDER BY Name")
        return Response(data)

#Container
class GetContainerFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 6 AND StatusID=1 ORDER BY Name")
        return Response(data)

#Making
class GetMakingFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId=12 AND StatusID=1 ORDER BY Name")
        return Response(data)


# Doubt need to check this code

# #VesselType
# class GetVesselTypeFromCommonMaster(APIView):
#     table = "CommonMaster"
#     def get(self, request):
#         data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId=17 AND StatusID=1 ORDER BY Name")
#         return Response(data)


#VesselType
class GetVesselTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId=14 AND StatusID=1 ORDER BY Name")
        return Response(data)

#CargoType
class GetCargoTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId=2 AND StatusID=1 ORDER BY Name")
        return Response(data)

#EngineCapacity
class GetEngineCapacityFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId=21 AND StatusID=1 ORDER BY Name")
        return Response(data)

#ReleaseLocation
class GetReleaseLocation(APIView):
    table = "ReleaseLocation"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * from {self.table} ORDER BY locationCode")
        return Response(data)

#ReceiptLocation
class GetReceiptLocation(APIView):
    table = "ReceiptLocation"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * from {self.table} ORDER BY locationCode")
        return Response(data)

#LoadingPort
class GetLoadingPort(APIView):
    table = "LoadingPort"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * from {self.table} ORDER BY portcode")
        return Response(data)

#StorageLocation
class GetStorageLocation(APIView):
    table = "StorageLocation"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * from {self.table} ORDER BY StorageCode")
        return Response(data)

#-----------------------------------------------------------------------------------------------#
#ManageUserMail

class GetManageUserMail(APIView):
    table = "ManageUser"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT DISTINCT MailBoxId  from {self.table} where Status='Active' ORDER BY MailBoxId ")
        return Response(data)

#-----------------------------------------------------------------------------------------------#
# Product Code By HSCode
class GetCascProductCodes(APIView):
    table = "CascProductCodes"
    def get(self, request):
        hscode = request.GET.get("HSCode")
        if hscode:
            query=f"""
                SELECT CASCCode, Description, UOM
                FROM {self.table}
                WHERE HSCode = '{hscode}'
                ORDER BY CASCCode
                """
        else:
            query=f"""
                SELECT CASCCode, Description, UOM
                FROM {self.table}
                ORDER BY CASCCode
                """
        data = SqlDb.execute_query(query)
        return Response(data)


#-----------------Excel Templates------------------------------
class ItemExcelDownload(APIView):
    def get(self, request, template_name):
        file_rel_path = settings.EXCEL_TEMPLATES.get(template_name)
        if not file_rel_path:
            return Response(
                {
                    "error": "Invalid template",
                    "available": list(settings.EXCEL_TEMPLATES.keys())
                },
                status=400
            )
        file_path = os.path.join(
            settings.EXCEL_TEMPLATE_DIR,
            file_rel_path
        )
        if not os.path.exists(file_path):
            return Response({"error": "File not found"}, status=404)
        return FileResponse(
            open(file_path, "rb"),
            as_attachment=True,
            filename=os.path.basename(file_path)
        )

def excel_to_date_str(val):
    if val is None or str(val).strip() == "" or str(val).strip() == "nan":
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    val_str = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y/%m/%d", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(val_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    try:
        serial = float(val_str)
        if serial > 0:
            excel_epoch = datetime(1899, 12, 30)
            converted = excel_epoch + timedelta(days=int(serial))
            return converted.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    return None

class ItemExcelUpload(APIView):
    def post(self, request):
        xlsx_file = request.FILES.get("file")
        print("Received file:", xlsx_file.name if xlsx_file else "No file")
        if not xlsx_file:
            return Response({"error": "No file provided"}, status=400)
        permit_id = request.POST.get("PermitId")
        msg_type = request.POST.get("MsgType")
        user_name = request.POST.get("UserName")
        touch_time = request.POST.get("TouchTime")
        # ---------------- READ EXCEL ----------------
        try:
            ItemInfo = pd.read_excel(
                xlsx_file, sheet_name="ItemInfo", dtype=str
            ).fillna('')
            CascInfo = pd.read_excel(
                xlsx_file, sheet_name="Casccodes", dtype=str
            ).fillna('')
        except Exception as e:
            return Response({"error": f"Failed to read Excel: {str(e)}"}, status=400)
        try:
            with connections['default'].cursor() as cursor:
                cursor.execute(
                    "SELECT COUNT(*) FROM CommonItemDtl WHERE PermitId=%s",
                    [permit_id]
                )
                existing_count = cursor.fetchone()[0]
        except Exception as e:
            return Response({"error": f"Failed to check existing items: {str(e)}"}, status=400)

        excel_count = len(ItemInfo)
        total_after_upload = existing_count + excel_count
        if total_after_upload > 150:
            return Response({
                "error": f"Upload blocked. You already have {existing_count} item(s). "
                         f"Excel contains {excel_count} item(s). "
                         f"Total {total_after_upload} exceeds the maximum limit of 150."
            }, status=400)

        ITEM_COLUMNS = [
            "ItemNo", "PermitId", "MessageType",
            "HSCode", "Description", "DGIndicator", "Contry",
            "EndUserDescription", "Brand", "Model",
            "InHAWBOBL", "OutHAWBOBL",
            "DutiableQty", "DutiableUOM",
            "TotalDutiableQty", "TotalDutiableUOM",
            "InvoiceQuantity", "HSQty", "HSUOM",
            "AlcoholPer", "InvoiceNo",
            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency",
            "ExchangeRate", "SumExchangeRate",
            "TotalLineAmount", "InvoiceCharges", "CIFFOB",
            "OPQty", "OPUOM",
            "IPQty", "IPUOM",
            "InPqty", "InPUOM",
            "ImPQty", "ImPUOM",
            "PreferentialCode",
            "GSTRate", "GSTUOM", "GSTAmount",
            "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
            "CurrentLot", "PreviousLot",
            "LSPValue", "Making",
            "ShippingMarks1", "ShippingMarks2",
            "ShippingMarks3", "ShippingMarks4",
            "CerItemQty", "CerItemUOM", "CIFValOfCer",
            "ManufactureCostDate", "TexCat",
            "TexQuotaQty", "TexQuotaUOM",
            "CerInvNo", "CerInvDate",
            "OriginOfCer", "HSCodeCer",
            "PerContent", "CertificateDescription",
            "TouchUser", "TouchTime",
            "VehicleType", "OptionalChrgeUOM",
            "EngineCapcity", "EngineCapUOM",
            "Optioncahrge", "OptionalSumtotal", "OptionalSumExchage",
            "orignaldatereg"
        ]

        CASC_COLUMNS = [
            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
            "RowNo", "CascCode1", "CascCode2", "CascCode3",
            "PermitId", "MessageType", "TouchUser", "TouchTime",
            "CASCId", "EndUserDes"
        ]

        ITEM_EXCEL_MAP = {
            "CountryofOrigin":        "Contry",
            "InvoiceNumber":          "InvoiceNo",
            "ItemCurrency":           "UnitPriceCurrency",
            "OuterPackQty":           "OPQty",
            "OuterPackUOM":           "OPUOM",
            "InPackQty":              "IPQty",
            "InPackUOM":              "IPUOM",
            "InnerPackQty":           "InPqty",
            "InnerPackUOM":           "InPUOM",
            "InmostPackQty":          "ImPQty",
            "InmostPackUOM":          "ImPUOM",
            "LastSellingPrice":       "LSPValue",
            "TarrifPreferentialCode": "PreferentialCode",
            "AlcoholPercentage":      "AlcoholPer",
        }

        try:
            with connections['default'].cursor() as cursor:
                item_no = 0
                for _, row in ItemInfo.iterrows():
                    item_no += 1
                    mapped = {}
                    for excel_col, val in row.items():
                        backend_col = ITEM_EXCEL_MAP.get(excel_col, excel_col)
                        mapped[backend_col] = val
                    values = [
                        item_no, permit_id, msg_type,
                        mapped.get("HSCode", ""),
                        mapped.get("Description", ""),
                        mapped.get("DGIndicator", ""),
                        mapped.get("Contry", ""),
                        mapped.get("EndUserDescription", ""),
                        mapped.get("Brand", ""),
                        mapped.get("Model", ""),
                        mapped.get("InHAWBOBL", ""),
                        mapped.get("OutHAWBOBL", ""),
                        to_float(mapped.get("DutiableQty")),
                        mapped.get("DutiableUOM", ""),
                        to_float(mapped.get("TotalDutiableQty")),
                        mapped.get("TotalDutiableUOM", ""),
                        to_float(mapped.get("InvoiceQuantity")),
                        to_float(mapped.get("HSQty")),
                        mapped.get("HSUOM", ""),
                        to_float(mapped.get("AlcoholPer")),
                        mapped.get("InvoiceNo", ""),
                        to_float(mapped.get("ChkUnitPrice")),
                        to_float(mapped.get("UnitPrice")),
                        mapped.get("UnitPriceCurrency", ""),
                        to_float(mapped.get("ExchangeRate")),
                        to_float(mapped.get("SumExchangeRate")),
                        to_float(mapped.get("TotalLineAmount")),
                        to_float(mapped.get("InvoiceCharges")),
                        to_float(mapped.get("CIFFOB")),
                        to_float(mapped.get("OPQty")),
                        mapped.get("OPUOM", ""),
                        to_float(mapped.get("IPQty")),
                        mapped.get("IPUOM", ""),
                        to_float(mapped.get("InPqty")),
                        mapped.get("InPUOM", ""),
                        to_float(mapped.get("ImPQty")),
                        mapped.get("ImPUOM", ""),
                        mapped.get("PreferentialCode", ""),
                        to_float(get_val(mapped, "GSTRate", 9)),
                        get_val(mapped, "GSTUOM", "PER"),
                        to_float(get_val(mapped, "GSTAmount", 0)),
                        to_float(mapped.get("ExciseDutyRate")),
                        mapped.get("ExciseDutyUOM", ""),
                        to_float(mapped.get("ExciseDutyAmount")),
                        to_float(mapped.get("CustomsDutyRate")),
                        mapped.get("CustomsDutyUOM", ""),
                        to_float(mapped.get("CustomsDutyAmount")),
                        to_float(mapped.get("OtherTaxRate")),
                        mapped.get("OtherTaxUOM", ""),
                        to_float(mapped.get("OtherTaxAmount")),
                        mapped.get("CurrentLot", ""),
                        mapped.get("PreviousLot", ""),
                        to_float(mapped.get("LSPValue")),
                        mapped.get("Making", ""),
                        mapped.get("ShippingMarks1", ""),
                        mapped.get("ShippingMarks2", ""),
                        mapped.get("ShippingMarks3", ""),
                        mapped.get("ShippingMarks4", ""),
                        to_float(mapped.get("CerItemQty")),
                        mapped.get("CerItemUOM", ""),
                        to_float(mapped.get("CIFValOfCer")),
                        excel_to_date_str(mapped.get("ManufactureCostDate")),
                        mapped.get("TexCat", ""),
                        to_float(mapped.get("TexQuotaQty")),
                        mapped.get("TexQuotaUOM", ""),
                        mapped.get("CerInvNo", ""),
                        excel_to_date_str(mapped.get("CerInvDate")),
                        mapped.get("OriginOfCer", ""),
                        mapped.get("HSCodeCer", ""),
                        mapped.get("PerContent", ""),
                        mapped.get("CertificateDescription", ""),
                        user_name, touch_time,
                        mapped.get("VehicleType", ""),
                        mapped.get("OptionalChrgeUOM", ""),
                        to_float(mapped.get("EngineCapcity")),
                        mapped.get("EngineCapUOM", ""),
                        to_float(mapped.get("Optioncahrge")),
                        to_float(mapped.get("OptionalSumtotal")),
                        to_float(mapped.get("OptionalSumExchage")),
                        mapped.get("orignaldatereg", "")
                    ]
                    cursor.execute(
                        "SELECT COUNT(*) FROM CommonItemDtl WHERE PermitId=%s AND ItemNo=%s",
                        [permit_id, item_no]
                    )
                    exists = cursor.fetchone()[0] > 0
                    if exists:
                        update_cols = ITEM_COLUMNS[3:]
                        set_clause = ", ".join([f"{col}=%s" for col in update_cols])
                        update_values = values[3:] + [permit_id, item_no]
                        cursor.execute(
                            f"UPDATE CommonItemDtl SET {set_clause} WHERE PermitId=%s AND ItemNo=%s",
                            update_values
                        )
                    else:
                        placeholders = ", ".join(["%s"] * len(ITEM_COLUMNS))
                        cursor.execute(
                            f"INSERT INTO CommonItemDtl ({','.join(ITEM_COLUMNS)}) VALUES ({placeholders})",
                            values
                        )
                # =========================================================
                # CASC INSERT / UPDATE
                # =========================================================
                for _, row in CascInfo.iterrows():
                    if not row.get("ProductCode"):
                        continue
                    casc_item_no = row.get("ItemNo", "")
                    casc_id = row.get("CASCId", "")
                    raw_row_no = row.get("RowNo", "")
                    row_no = int(raw_row_no) if str(raw_row_no).strip() != '' else 1
                    cursor.execute(
                        """
                        SELECT COUNT(*) FROM CommonCASCDtl
                        WHERE ItemNo=%s AND PermitId=%s AND RowNo=%s AND CASCId=%s
                        """,
                        [casc_item_no, permit_id, row_no, casc_id]
                    )
                    casc_exists = cursor.fetchone()[0] > 0
                    if casc_exists:
                        cursor.execute(
                            """
                            UPDATE CommonCASCDtl
                            SET ProductCode=%s, Quantity=%s, ProductUOM=%s,
                                CascCode1=%s, CascCode2=%s, CascCode3=%s,
                                TouchUser=%s, TouchTime=%s, EndUserDes=%s
                            WHERE ItemNo=%s AND PermitId=%s AND RowNo=%s AND CASCId=%s
                            """,
                            [
                                row.get("ProductCode"),
                                to_float(row.get("Quantity")),
                                row.get("ProductUOM"),
                                row.get("CascCode1", ""),
                                row.get("CascCode2", ""),
                                row.get("CascCode3", ""),
                                user_name, touch_time,
                                row.get("EndUserDes", ""),
                                casc_item_no, permit_id, row_no, casc_id
                            ]
                        )
                    else:
                        cursor.execute(
                            f"""
                            INSERT INTO CommonCASCDtl ({','.join(CASC_COLUMNS)})
                            VALUES ({','.join(['%s'] * len(CASC_COLUMNS))})
                            """,
                            [
                                casc_item_no,
                                row.get("ProductCode"),
                                to_float(row.get("Quantity")),
                                row.get("ProductUOM"),
                                row_no,
                                row.get("CascCode1", ""),
                                row.get("CascCode2", ""),
                                row.get("CascCode3", ""),
                                permit_id, msg_type,
                                user_name, touch_time,
                                casc_id,
                                row.get("EndUserDes", "")
                            ]
                        )
                connections['default'].commit()
        except Exception as e:
            return Response({"error": f"Upload failed: {str(e)}"}, status=400)
        # =========================================================
        # FETCH & RETURN
        # =========================================================
        try:
            with connections['default'].cursor() as cursor:
                cursor.execute(
                    "SELECT * FROM CommonItemDtl WHERE PermitId=%s ORDER BY ItemNo",
                    [permit_id]
                )
                item_columns = [col[0] for col in cursor.description]
                items = [dict(zip(item_columns, r)) for r in cursor.fetchall()]
                cursor.execute(
                    "SELECT * FROM CommonCASCDtl WHERE PermitId=%s ORDER BY ItemNo",
                    [permit_id]
                )
                casc_columns = [col[0] for col in cursor.description]
                casc = [dict(zip(casc_columns, r)) for r in cursor.fetchall()]
        except Exception as e:
            return Response({"error": f"Fetch failed: {str(e)}"}, status=400)
        return Response({
            "Result": "UPLOAD SUCCESSFULLY",
            "item": items,
            "casc": casc
        }, status=201)


#Edit all Item
class AllItemUpdate(APIView):
    def post(self, request):
        try:
            Item = request.data.get("Item")
            PermitId = request.data.get("PermitId")

            if not Item or not PermitId:
                return Response({"message": "Item and PermitId are required"}, status=400)

            for item in Item:
                SqlDb.execute_query("""
                    UPDATE CommonItemDtl SET
                        MessageType=%s, HSCode=%s, Description=%s,
                        DGIndicator=%s, Contry=%s, Brand=%s, Model=%s,
                        InHAWBOBL=%s,OutHAWBOBL=%s, DutiableQty=%s, DutiableUOM=%s,
                        TotalDutiableQty=%s, TotalDutiableUOM=%s,
                        InvoiceQuantity=%s, HSQty=%s, HSUOM=%s,
                        AlcoholPer=%s, InvoiceNo=%s, ChkUnitPrice=%s,
                        UnitPrice=%s, UnitPriceCurrency=%s, ExchangeRate=%s,
                        SumExchangeRate=%s, TotalLineAmount=%s,
                        InvoiceCharges=%s, CIFFOB=%s,
                        OPQty=%s, OPUOM=%s, IPQty=%s, IPUOM=%s,
                        InPqty=%s, InPUOM=%s, ImPQty=%s, ImPUOM=%s,
                        PreferentialCode=%s, GSTRate=%s, GSTUOM=%s,
                        GSTAmount=%s, ExciseDutyRate=%s, ExciseDutyUOM=%s,
                        ExciseDutyAmount=%s, CustomsDutyRate=%s,
                        CustomsDutyUOM=%s, CustomsDutyAmount=%s,
                        OtherTaxRate=%s, OtherTaxUOM=%s, OtherTaxAmount=%s,
                        CurrentLot=%s, PreviousLot=%s, LSPValue=%s,
                        Making=%s, ShippingMarks1=%s, ShippingMarks2=%s,
                        ShippingMarks3=%s, ShippingMarks4=%s,
                        TouchUser=%s, TouchTime=%s, VehicleType=%s,
                        EngineCapcity=%s, EngineCapUOM=%s,
                        orignaldatereg=%s, OptionalChrgeUOM=%s,
                        Optioncahrge=%s, OptionalSumtotal=%s,
                        OptionalSumExchage=%s
                    WHERE ItemNo=%s AND PermitId=%s
                """, [
                    item.get("MessageType"),
                    item.get("HSCode"),
                    item.get("Description"),
                    item.get("DGIndicator"),
                    item.get("Contry"),
                    item.get("Brand"),
                    item.get("Model"),
                    item.get("InHAWBOBL"),
                    item.get("OutHAWBOBL"),
                    item.get("DutiableQty"),
                    item.get("DutiableUOM"),
                    item.get("TotalDutiableQty"),
                    item.get("TotalDutiableUOM"),
                    item.get("InvoiceQuantity"),
                    item.get("HSQty"),
                    item.get("HSUOM"),
                    item.get("AlcoholPer"),
                    item.get("InvoiceNo"),
                    item.get("ChkUnitPrice"),
                    item.get("UnitPrice"),
                    item.get("UnitPriceCurrency"),
                    item.get("ExchangeRate"),
                    item.get("SumExchangeRate"),
                    item.get("TotalLineAmount"),
                    item.get("InvoiceCharges"),
                    item.get("CIFFOB"),
                    item.get("OPQty"),
                    item.get("OPUOM"),
                    item.get("IPQty"),
                    item.get("IPUOM"),
                    item.get("InPqty"),
                    item.get("InPUOM"),
                    item.get("ImPQty"),
                    item.get("ImPUOM"),
                    item.get("PreferentialCode"),
                    item.get("GSTRate"),
                    item.get("GSTUOM"),
                    item.get("GSTAmount"),
                    item.get("ExciseDutyRate"),
                    item.get("ExciseDutyUOM"),
                    item.get("ExciseDutyAmount"),
                    item.get("CustomsDutyRate"),
                    item.get("CustomsDutyUOM"),
                    item.get("CustomsDutyAmount"),
                    item.get("OtherTaxRate"),
                    item.get("OtherTaxUOM"),
                    item.get("OtherTaxAmount"),
                    item.get("CurrentLot"),
                    item.get("PreviousLot"),
                    item.get("LSPValue"),
                    item.get("Making"),
                    item.get("ShippingMarks1"),
                    item.get("ShippingMarks2"),
                    item.get("ShippingMarks3"),
                    item.get("ShippingMarks4"),
                    item.get("TouchUser"),
                    item.get("TouchTime"),
                    item.get("VehicleType"),
                    item.get("EngineCapcity"),
                    item.get("EngineCapUOM"),
                    item.get("orignaldatereg"),
                    item.get("OptionalChrgeUOM"),
                    item.get("Optioncahrge"),
                    item.get("OptionalSumtotal"),
                    item.get("OptionalSumExchage"),
                    item.get("ItemNo"),
                    PermitId,
                ])


            items = SqlDb.execute_query(
                "SELECT * FROM CommonItemDtl WHERE PermitId=%s ORDER BY ItemNo",
                (PermitId,)
            )
            casc = SqlDb.execute_query(
                "SELECT * FROM CommonCASCDtl WHERE PermitId=%s ORDER BY ItemNo",
                (PermitId,)
            )

            return Response({
                "Item": items,
                "ItemCasc": casc,
                "message": "All Item Updated Successfully!"
            }, status=200)

        except Exception as e:
            print(e)
            return Response({
                "message": f"Update failed: {str(e)}"
            }, status=400)
class UpdateItemBrandByPermitId(APIView):
    table = "CommonItemDtl"

    def post(self, request):
        try:
            permit_id = request.data.get("PermitId")
            brand = request.data.get("Brand", "")

            if not permit_id:
                return Response({"error": "PermitId is required"}, status=400)

            SqlDb.execute_query(
                f"UPDATE {self.table} SET Brand=%s WHERE PermitId=%s",
                [brand, permit_id],
            )
            SqlDb.commit()

            records = SqlDb.execute_query(
                f"SELECT * FROM {self.table} WHERE PermitId=%s ORDER BY ItemNo",
                [permit_id],
            )

            return Response(
                {
                    "message": (
                        f"Brand updated to '{brand}' for all items"
                        if brand
                        else "Brand cleared for all items"
                    ),
                    "Records": records,
                },
                status=200,
            )
        except Exception as e:
            return Response({"error": f"Failed to update brand: {str(e)}"}, status=400)
class DeleteHawbl(APIView):
    def delete(self, request, permit_id):
        try:
            SqlDb.execute_query(
                "UPDATE CommonItemDtl SET InHAWBOBL='', OutHAWBOBL='' WHERE PermitId=%s",
                (permit_id,)
            )
            return Response(
                {"message": f"HAWBOBL cleared for all items in PermitId {permit_id}"},
                status=200
            )
        except Exception as e:
            return Response(
                {"error": f"Failed to clear HAWBOBL: {str(e)}"},
                status=400
            )

# New Permit
class InpaymentNewPermit(APIView):
    def get(self, request):
        try:
            Username = request.query_params.get("user")
            if not Username:
                Username = request.session.get("Username")
            if not Username:
                return Response({"error": "Session expired or User not provided"}, status=401)
            refDate = datetime.now().strftime("%Y%m%d")
            yy_mmdd = datetime.now().strftime("%Y%m%d")
            currentDate = datetime.now().strftime("%d/%m/%Y")  

            q_account = "SELECT AccountId FROM ManageUser WHERE UserName = %s"

            account_rows = SqlDb.execute_query(q_account, [Username])

            if not account_rows:
                return Response({"error": "User not found"}, status=404)
            AccountId = account_rows[0]['AccountId']

            # q_ref = """
            #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
            #     FROM CommonHeaderTbl 
            #     WHERE MSGId LIKE %s AND MessageType = 'IPTDEC'
            # """
            # ref_rows = SqlDb.execute_query(q_ref, [f"%{refDate}%"])
            # ref_rows = SqlDb.execute_query(q_ref, [f"{refDate}%"])

            # ref_rows = SqlDb.execute_query(
            #     """
            #     SELECT ISNULL(COUNT(*), 0) + 1 AS Count
            #     FROM CommonHeaderTbl
            #     WHERE PermitId LIKE %s
            #     """,
            #     [f"{Username}{refDate}%"]
            # )

            # RefId = "%03d" % (ref_rows[0]['Count'] if ref_rows else 1)

            # q_job = """
            #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
            #     FROM PermitCount 
            #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'IPTDEC'
            # """

            # job_rows = SqlDb.execute_query(
                # """
                # # SELECT ISNULL(MAX(CAST(SUBSTRING(JobId, 2, 6) + RIGHT(JobId, 5) AS BIGINT)), 0) + 1 AS Count
                # # FROM CommonHeaderTbl
                # """
            # job_rows = SqlDb.execute_query(
            #     """
            #     SELECT ISNULL(COUNT(*),0) + 1 as Count 
            #     FROM PermitCount 
            #     WHERE TouchTime LIKE %s AND AccountId = %s AND MessageType = 'IPTDEC'
            #     """,
            #     [f"{jobDate}%", AccountId] 
            # )

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
            #     WHERE TouchTime LIKE %s AND AccountId = %s
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


# class InpaymentList(APIView):
#     def get(self, request):
#         try:
#             Username = request.query_params.get("user") or request.session.get("Username")
#             if not Username:
#                 return Response({"error": "Username required"}, status=400)
            
#             q_account = "SELECT AccountId FROM ManageUser WHERE UserName = %s"
#             account_rows = SqlDb.execute_query(q_account, [Username])
#             if not account_rows:
#                 return Response({"error": "User not found"}, status=404)
#             AccountId = account_rows[0]["AccountId"]
            
#             show_all = request.query_params.get("all", "false").lower() == "true"

#             if show_all:
#                 query = """
#                     SELECT  
#                         t1.Id AS ID,
#                         t1.JobId,
#                         t1.PermitId,
#                         t1.MSGId,
#                         CONVERT(varchar, t1.TouchTime, 105) AS DECDATE,
#                         SUBSTRING(t1.DeclarationType, 1, CHARINDEX(':', t1.DeclarationType) - 1) AS DECTYPE,
#                         t1.TouchUser AS CREATE_USER,
#                         u.MailBoxId AS DECID,
#                         CONVERT(varchar, t1.ArrivalDate, 105) AS ETA,
#                         t1.PermitNumber AS PERMITNO,
#                         i.Name + ' ' + i.Name1 AS IMPORTER,
#                         t1.HBL AS HAWB,
#                         CASE  
#                             WHEN t1.InwardTransportMode = '4 : Air' THEN t1.MasterAirwayBill  
#                             WHEN t1.InwardTransportMode = '1 : Sea' THEN t1.OceanBillofLadingNo  
#                             ELSE '' 
#                         END AS MAWBOBL,
#                         t1.LoadingPortCode AS POL,
#                         t1.MessageType AS MSGTYPE,
#                         t1.InwardTransportMode AS TPT,
#                         t1.PreviousPermit AS PREPMT,
#                         t1.GrossReference AS XREF,
#                         t1.InternalRemarks AS INTREM,
#                         t1.TotalGSTTaxAmt AS GSTAMT,
#                         t1.Status,
#                         CASE  
#                             WHEN t1.Status = 'APR' THEN 
#                                 CASE  
#                                     WHEN EXISTS (SELECT 1 FROM CommonPMT p WHERE p.PermitNumber = t1.PermitNumber AND p.ConditionCode IN ('Z02','Z18','Z06')) THEN 'RED'
#                                     WHEN EXISTS (SELECT 1 FROM CommonPMT p WHERE p.PermitNumber = t1.PermitNumber AND p.ConditionCode IN ('D6','D3')) THEN 'MAROON'
#                                     ELSE 'DEFAULT'
#                                 END
#                             ELSE 'DEFAULT'
#                         END AS COLOR
#                     FROM CommonHeaderTbl t1
#                     LEFT JOIN ManageUser u ON u.UserName = t1.TouchUser
#                     LEFT JOIN CommonImporter i ON t1.ImporterCompanyCode = i.Code
#                     WHERE u.AccountId = %s
#                     AND t1.MessageType = 'IPTDEC'
#                     ORDER BY t1.Id DESC
#                 """
#                 result = SqlDb.execute_query(query, [AccountId])  # ← only AccountId

#             else:
#                 nowdate = datetime.now() - timedelta(days=90)
#                 date_filter = nowdate.strftime("%Y/%m/%d")

#                 query = """
#                     SELECT  
#                         t1.Id AS ID,
#                         t1.JobId,
#                         t1.PermitId,
#                         t1.MSGId,
#                         CONVERT(varchar, t1.TouchTime, 105) AS DECDATE,
#                         SUBSTRING(t1.DeclarationType, 1, CHARINDEX(':', t1.DeclarationType) - 1) AS DECTYPE,
#                         t1.TouchUser AS CREATE_USER,
#                         u.MailBoxId AS DECID,
#                         CONVERT(varchar, t1.ArrivalDate, 105) AS ETA,
#                         t1.PermitNumber AS PERMITNO,
#                         i.Name + ' ' + i.Name1 AS IMPORTER,
#                         t1.HBL AS HAWB,
#                         CASE  
#                             WHEN t1.InwardTransportMode = '4 : Air' THEN t1.MasterAirwayBill  
#                             WHEN t1.InwardTransportMode = '1 : Sea' THEN t1.OceanBillofLadingNo  
#                             ELSE '' 
#                         END AS MAWBOBL,
#                         t1.LoadingPortCode AS POL,
#                         t1.MessageType AS MSGTYPE,
#                         t1.InwardTransportMode AS TPT,
#                         t1.PreviousPermit AS PREPMT,
#                         t1.GrossReference AS XREF,
#                         t1.InternalRemarks AS INTREM,
#                         t1.TotalGSTTaxAmt AS GSTAMT,
#                         t1.Status,
#                         CASE  
#                             WHEN t1.Status = 'APR' THEN 
#                                 CASE  
#                                     WHEN EXISTS (SELECT 1 FROM CommonPMT p WHERE p.PermitNumber = t1.PermitNumber AND p.ConditionCode IN ('Z02','Z18','Z06')) THEN 'RED'
#                                     WHEN EXISTS (SELECT 1 FROM CommonPMT p WHERE p.PermitNumber = t1.PermitNumber AND p.ConditionCode IN ('D6','D3')) THEN 'MAROON'
#                                     ELSE 'DEFAULT'
#                                 END
#                             ELSE 'DEFAULT'
#                         END AS COLOR
#                     FROM CommonHeaderTbl t1
#                     LEFT JOIN ManageUser u ON u.UserName = t1.TouchUser
#                     LEFT JOIN CommonImporter i ON t1.ImporterCompanyCode = i.Code
#                     WHERE u.AccountId = %s
#                         AND t1.MessageType = 'IPTDEC'
#                         AND CONVERT(varchar, t1.TouchTime, 111) >= %s
#                     ORDER BY t1.Id DESC
#                 """
#                 result = SqlDb.execute_query(query, [AccountId, date_filter])  # ← both params

#             return Response(result)

#         except Exception as e:
#             traceback.print_exc()
#             return Response({"error": str(e)}, status=500)

class InpaymentList(APIView):
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
                    AND t1.MessageType = 'IPTDEC'
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId])

            else:
                nowdate = datetime.now() - timedelta(days=90)
                date_filter = nowdate.strftime("%Y/%m/%d")

                query = base_select + """
                    WHERE t1.TradeNetMailboxID = %s
                    AND t1.MessageType = 'IPTDEC'
                    AND CONVERT(varchar, t1.TouchTime, 111) >= %s
                    ORDER BY t1.Id DESC
                """
                result = SqlDb.execute_query(query, [MailBoxId, date_filter])

            return Response(result)

        except Exception as e:
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)

# class CopyInpayment(APIView):
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

#                     # RefId — count IPTDEC records for today
#                     # cursor.execute("""
#                     #     SELECT ISNULL(COUNT(*), 0) + 1
#                     #     FROM CommonHeaderTbl
#                     #     WHERE MSGId LIKE %s AND MessageType = 'IPTDEC'
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
#                         VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
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


class GetDeclarantByMailbox(APIView):
    def get(self, request):
        mailbox_id = request.GET.get("MailboxId")
        if not mailbox_id:
            return Response({"error": "MailboxId is required"}, status=400)
        rows = SqlDb.execute_query(
            """SELECT Code, CRUEI, DeclarantName, DeclarantCode, DeclarantTel, Name, Name1
               FROM DeclarantCompany WHERE TradeNetMailboxID = %s""",
            [mailbox_id]
        )
        if not rows:
            return Response({"error": "Declarant not found"}, status=404)
        return Response(rows[0])


# class CopyInpayment(APIView):
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

#                 # GLOBAL starting count for JobId/MsgId (matches PostCommonHeaderTable logic)
#                 cursor.execute("""
#                     SELECT ISNULL(COUNT(*), 0) + 1 AS Count
#                     FROM CommonHeaderTbl
#                     WHERE JobId LIKE %s
#                 """, [f"K{job_date}%"])
#                 job_count = cursor.fetchone()[0]

#                 # GLOBAL starting count for PermitId/RefId (per-user, today)
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
#                         VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
#                     """, [new_permit_id, account_id, msg_id, username, touch_time])

#                     copied_permits.append(new_permit_id)

#                     # Increment counters for next permit in this batch
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


class CopyInpayment(APIView):

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

                    # ── 1. Insert new CommonHeaderTbl row (unchanged from before) ──
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
                            %s, %s, %s, %s, %s,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
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
                            Defrentprinting, Cnb, NULL, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. NEW: mirror the row we just inserted into InHeaderTbl ──
                    # Reads from the NEW CommonHeaderTbl row (new_permit_id), so no
                    # recomputation of Ref/Job/Msg ids or dates — single source of truth.
                    cursor.execute("""
                        INSERT INTO InHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ReleaseLocName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
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
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: copy into BOTH Common* and In-payment* targets ──
                    # Each entry: source_table -> (common_target, in_target, cols)
                    child_tables = {
                        "CommonInvoiceDtl": ("InvoiceDtl", [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
                        ]),
                        "CommonItemDtl": ("ItemDtl", [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "Brand", "Model", "InHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode", "GSTRate", "GSTUOM",
                            "GSTAmount", "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "LSPValue", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                        ]),
                        "CommonCASCDtl": ("CASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM", "RowNo",
                            "CascCode1", "CascCode2", "CascCode3", "MessageType",
                            "TouchUser", "TouchTime", "CASCId"
                        ]),
                        "CommonCPCDtl": ("CPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("ContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, (dst_table, cols) in child_tables.items():
                        col_list = ", ".join(cols)
                        # Copy Common -> Common (from OLD permit, as before)
                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")

                        # Copy Common(NEW) -> In-payment target — single read of what we
                        # just wrote, so both tables end up byte-identical for this permit.
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # File & PMT — column names differ, handle explicitly
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno,Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                        cursor.execute("""
                            INSERT INTO InFile (PermitId,Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                                SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
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
                        # InPMT has a different schema (Sno, CertificateNumber, StartDate, etc.)
                        # — usually populated on approval, not at copy time. Skip unless you
                        # need to carry PMT conditions forward too; adjust columns if so.
                    except Exception as err:
                        print(f"Warning copying PMT: {err}")

                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
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

# Print Gst
class PrintGst(APIView):
    def get(self, request, PermitId):
        try:
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT JobId, MSGId, PermitNumber, TouchTime
                    FROM CommonHeaderTbl
                    WHERE PermitId = %s
                """, [PermitId])
                header_row = cursor.fetchone()
            if not header_row:
                return HttpResponse("Permit not found", status=404)
            JobId, MSGId, PermitNumber, TouchTime = header_row

            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT
                        InvoiceNo, TermType, ImportPartyCode,
                        TICurrency, TIExRate, TIAmount, TISAmount,
                        FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                        ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                        OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                        CIFSUMAmount, GSTPercentage, GSTSUMAmount
                    FROM CommonInvoiceDtl
                    WHERE PermitId = %s
                """, [PermitId])
                invoices = cursor.fetchall()

            if not invoices:
                return HttpResponse("No invoice details found", status=404)

            import_party_code = invoices[0][2]
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT Name, CRUEI
                    FROM CommonImporter
                    WHERE Code = %s
                """, [import_party_code])
                importer_row = cursor.fetchone()

            if not importer_row:
                return HttpResponse("Importer not found", status=404)
            ImportName, ImportCRUEI = importer_row

            # ── BUILD PDF ─────────────────────────────────────────────
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(595, 841))

            y = 820
            y = self._draw_header(
                can, y, ImportName, ImportCRUEI,
                JobId, MSGId, PermitNumber, TouchTime
            )

            for invoice in invoices:
                y = self._draw_invoice(
                    can, y, invoice,
                    ImportName, ImportCRUEI,
                    JobId, MSGId, PermitNumber, TouchTime
                )

            can.save()
            packet.seek(0)

            response = HttpResponse(packet, content_type='application/pdf')
            response['Content-Disposition'] = (
                f'attachment; filename="{PermitNumber}_GST.pdf"'
            )
            return response

        except Exception as e:
            import traceback
            traceback.print_exc()
            return HttpResponse(f"Error generating PDF: {str(e)}", status=500)

    # ─────────────────────────────────────────────────────────────────
    def _safe_val(self, val):
        if val is None:
            return ""
        if str(val) == '--Select--':
            return ""
        return str(val)

    # ─────────────────────────────────────────────────────────────────
    def _draw_header(self, can, y, ImportName, ImportCRUEI,
                     JobId, MSGId, PermitNumber, TouchTime):
        can.setFont('Times-Bold', 9)
        can.drawString(150, y, "GOODS AND SERVICE TAX (GST) - CALCULATION SHEET")

        y -= 20
        can.drawString(10,  y, "COMPANY NAME:")
        can.drawString(440, y, "JOB NUMBER:")
        y -= 15
        can.drawString(10,  y, "COMPANY UEN:")
        can.drawString(440, y, "JOB CREATED:")
        y -= 15
        can.drawString(10,  y, "PERMIT NUMBER:")
        can.drawString(440, y, "MESSAGE ID:")

        can.setFont('Times-Roman', 9)
        y += 30
        can.drawString(110, y, str(ImportName or ""))
        can.drawString(510, y, str(JobId or ""))
        y -= 15
        can.drawString(110, y, str(ImportCRUEI or ""))
        can.drawString(
            510, y,
            TouchTime.strftime("%d/%m/%Y  %H:%M:%S") if TouchTime else ""
        )
        y -= 15
        can.drawString(110, y, str(PermitNumber or ""))
        can.drawString(510, y, str(MSGId or ""))

        return y

    # ─────────────────────────────────────────────────────────────────
    def _draw_invoice(self, can, y, invoice,
                      ImportName, ImportCRUEI,
                      JobId, MSGId, PermitNumber, TouchTime):
        (
            InvoiceNo, TermType, ImportPartyCode,
            TICurrency, TIExRate, TIAmount, TISAmount,
            FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
            ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
            OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
            CIFSUMAmount, GSTPercentage, GSTSUMAmount
        ) = invoice

        can.setFont('Times-Bold', 9)
        y -= 20
        can.drawString(10, y, "-" * 165)

        y -= 15
        can.drawString(10, y, "INVOICE NUMBER:")
        can.setFont('Times-Roman', 9)
        can.drawString(110, y, self._safe_val(InvoiceNo))

        y -= 15
        can.setFont('Times-Bold', 9)
        can.drawString(10, y, "INVOICE TERM:")
        can.setFont('Times-Roman', 9)
        can.drawString(110, y, self._safe_val(TermType))

        y -= 10

        table_data = [
            ['SNO', 'ITEM', 'CURRENCY', 'EXCHG. RATE', 'PERC.', 'AMOUNT', 'AMOUNT (SGD)'],
            ["1", "TOTAL INVOICE",
             self._safe_val(TICurrency), self._safe_val(TIExRate),
             "", self._safe_val(TIAmount), self._safe_val(TISAmount)],
            ["2", "FREIGHT CHARGES",
             self._safe_val(FCCurrency), self._safe_val(FCExRate),
             self._safe_val(FCCharge), self._safe_val(FCAmount), self._safe_val(FCSAmount)],
            ["3", "INSURANCE",
             self._safe_val(ICCurrency), self._safe_val(ICExRate),
             self._safe_val(ICCharge), self._safe_val(ICAmount), self._safe_val(ICSAmount)],
            ["4", "OTHER CHARGES",
             self._safe_val(OTCCurrency), self._safe_val(OTCExRate),
             self._safe_val(OTCCharge), self._safe_val(OTCAmount), self._safe_val(OTCSAmount)],
            ["5", "CUSTOMS VALUE",
             "", "", "", "", self._safe_val(CIFSUMAmount)],
            ["6", "GST",
             "", "", self._safe_val(GSTPercentage), "", self._safe_val(GSTSUMAmount)],
        ]

        col_widths = [30, 110, 75, 75, 55, 80, 95]
        table = Table(table_data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('ALIGN',          (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME',       (0, 0), (-1,  0), 'Times-Bold'),
            ('FONTNAME',       (0, 1), (-1, -1), 'Times-Roman'),
            ('FONTSIZE',       (0, 0), (-1, -1), 9),
            ('GRID',           (0, 0), (-1, -1), 0.5, colors.black),
            ('BOX',            (0, 0), (-1, -1), 1,   colors.black),
            ('VALIGN',         (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 0), (-1,  0), [colors.lightgrey]),
        ]))

        row_height   = 18
        table_height = row_height * len(table_data)

        # ── Page break if not enough space ──
        if y - table_height < 60:
            can.showPage()
            y = 820
            y = self._draw_header(
                can, y, ImportName, ImportCRUEI,
                JobId, MSGId, PermitNumber, TouchTime
            )
            y -= 20

        table.wrapOn(can, 0, 0)
        table.drawOn(can, 10, y - table_height)
        y -= (table_height + 15)

        return y

# print Gst all
class PrintGstAll(APIView):
    def get(self, request):
        data = request.GET.get("data", "")
        if not data:
            return HttpResponse("No permits provided", status=400)

        pdf_files = []
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        PDF_DIR = os.path.join(BASE_DIR, "PrintGstAll")
        os.makedirs(PDF_DIR, exist_ok=True)

        for permit_id in data.split(','):
            permit_id = permit_id.strip()
            if not permit_id:
                continue

            try:
                # ── Fetch header ───────────────────────────────────────
                with connection.cursor() as cursor:
                    cursor.execute("""
                        SELECT JobId, MSGId, PermitNumber, TouchTime
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [permit_id])
                    header_row = cursor.fetchone()

                if not header_row:
                    continue

                JobId, MSGId, PermitNumber, TouchTime = header_row

                # ── Fetch invoices ─────────────────────────────────────
                with connection.cursor() as cursor:
                    cursor.execute("""
                        SELECT
                            InvoiceNo, TermType, ImportPartyCode,
                            TICurrency, TIExRate, TIAmount, TISAmount,
                            FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                            ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                            OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                            CIFSUMAmount, GSTPercentage, GSTSUMAmount
                        FROM CommonInvoiceDtl
                        WHERE PermitId = %s
                    """, [permit_id])
                    invoices = cursor.fetchall()

                if not invoices:
                    continue

                # ── Fetch importer ─────────────────────────────────────
                import_party_code = invoices[0][2]
                with connection.cursor() as cursor:
                    cursor.execute("""
                        SELECT Name, CRUEI
                        FROM CommonImporter
                        WHERE Code = %s
                    """, [import_party_code])
                    importer_row = cursor.fetchone()

                if not importer_row:
                    continue

                ImportName, ImportCRUEI = importer_row

                # ── Build PDF into buffer ──────────────────────────────
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(595, 841))

                y = 820
                y = self._draw_header(
                    can, y, ImportName, ImportCRUEI,
                    JobId, MSGId, PermitNumber, TouchTime
                )

                for invoice in invoices:
                    y = self._draw_invoice(
                        can, y, invoice,
                        ImportName, ImportCRUEI,
                        JobId, MSGId, PermitNumber, TouchTime
                    )

                can.save()
                packet.seek(0)

                # ── Save PDF to disk ───────────────────────────────────
                safe_name = (PermitNumber or permit_id).strip().replace(" ", "_")
                pdf_path = os.path.join(PDF_DIR, f"{safe_name}_GST.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(packet.read())
                packet.close()

                pdf_files.append(pdf_path)

            except Exception as e:
                import traceback
                traceback.print_exc()
                continue

        if not pdf_files:
            return HttpResponse("No PDFs generated", status=400)

        # ── Zip all PDFs ───────────────────────────────────────────────
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, mode='w') as zf:
            for fpath in pdf_files:
                if os.path.isfile(fpath):
                    zf.write(fpath, arcname=os.path.join("GST_Files", os.path.basename(fpath)))

        # ── Clean up temp PDFs ─────────────────────────────────────────
        for fpath in pdf_files:
            if os.path.isfile(fpath):
                try:
                    os.remove(fpath)
                except:
                    pass

        zip_buffer.seek(0)
        response = HttpResponse(zip_buffer, content_type='application/zip')
        response['Content-Disposition'] = (
            f'attachment; filename="GST_{datetime.now().strftime("%Y-%B-%d-%H%M")}.zip"'
        )
        return response

    # ── Reuse exact same helpers from PrintGst ─────────────────────────
    def _safe_val(self, val):
        if val is None:
            return ""
        if str(val) == '--Select--':
            return ""
        return str(val)

    def _draw_header(self, can, y, ImportName, ImportCRUEI,
                     JobId, MSGId, PermitNumber, TouchTime):
        can.setFont('Times-Bold', 9)
        can.drawString(150, y, "GOODS AND SERVICE TAX (GST) - CALCULATION SHEET")

        y -= 20
        can.drawString(10,  y, "COMPANY NAME:")
        can.drawString(440, y, "JOB NUMBER:")
        y -= 15
        can.drawString(10,  y, "COMPANY UEN:")
        can.drawString(440, y, "JOB CREATED:")
        y -= 15
        can.drawString(10,  y, "PERMIT NUMBER:")
        can.drawString(440, y, "MESSAGE ID:")

        can.setFont('Times-Roman', 9)
        y += 30
        can.drawString(110, y, str(ImportName or ""))
        can.drawString(510, y, str(JobId or ""))
        y -= 15
        can.drawString(110, y, str(ImportCRUEI or ""))
        can.drawString(
            510, y,
            TouchTime.strftime("%d/%m/%Y  %H:%M:%S") if TouchTime else ""
        )
        y -= 15
        can.drawString(110, y, str(PermitNumber or ""))
        can.drawString(510, y, str(MSGId or ""))

        return y

    def _draw_invoice(self, can, y, invoice,
                      ImportName, ImportCRUEI,
                      JobId, MSGId, PermitNumber, TouchTime):
        from reportlab.platypus import Table, TableStyle
        from reportlab.lib import colors

        (
            InvoiceNo, TermType, ImportPartyCode,
            TICurrency, TIExRate, TIAmount, TISAmount,
            FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
            ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
            OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
            CIFSUMAmount, GSTPercentage, GSTSUMAmount
        ) = invoice

        can.setFont('Times-Bold', 9)
        y -= 20
        can.drawString(10, y, "-" * 165)

        y -= 15
        can.drawString(10, y, "INVOICE NUMBER:")
        can.setFont('Times-Roman', 9)
        can.drawString(110, y, self._safe_val(InvoiceNo))

        y -= 15
        can.setFont('Times-Bold', 9)
        can.drawString(10, y, "INVOICE TERM:")
        can.setFont('Times-Roman', 9)
        can.drawString(110, y, self._safe_val(TermType))

        y -= 10

        table_data = [
            ['SNO', 'ITEM', 'CURRENCY', 'EXCHG. RATE', 'PERC.', 'AMOUNT', 'AMOUNT (SGD)'],
            ["1", "TOTAL INVOICE",
             self._safe_val(TICurrency), self._safe_val(TIExRate),
             "", self._safe_val(TIAmount), self._safe_val(TISAmount)],
            ["2", "FREIGHT CHARGES",
             self._safe_val(FCCurrency), self._safe_val(FCExRate),
             self._safe_val(FCCharge), self._safe_val(FCAmount), self._safe_val(FCSAmount)],
            ["3", "INSURANCE",
             self._safe_val(ICCurrency), self._safe_val(ICExRate),
             self._safe_val(ICCharge), self._safe_val(ICAmount), self._safe_val(ICSAmount)],
            ["4", "OTHER CHARGES",
             self._safe_val(OTCCurrency), self._safe_val(OTCExRate),
             self._safe_val(OTCCharge), self._safe_val(OTCAmount), self._safe_val(OTCSAmount)],
            ["5", "CUSTOMS VALUE",
             "", "", "", "", self._safe_val(CIFSUMAmount)],
            ["6", "GST",
             "", "", self._safe_val(GSTPercentage), "", self._safe_val(GSTSUMAmount)],
        ]

        col_widths = [30, 110, 75, 75, 55, 80, 95]
        table = Table(table_data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('ALIGN',          (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME',       (0, 0), (-1,  0), 'Times-Bold'),
            ('FONTNAME',       (0, 1), (-1, -1), 'Times-Roman'),
            ('FONTSIZE',       (0, 0), (-1, -1), 9),
            ('GRID',           (0, 0), (-1, -1), 0.5, colors.black),
            ('BOX',            (0, 0), (-1, -1), 1,   colors.black),
            ('VALIGN',         (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 0), (-1,  0), [colors.lightgrey]),
        ]))

        row_height   = 18
        table_height = row_height * len(table_data)

        if y - table_height < 60:
            can.showPage()
            y = 820
            y = self._draw_header(
                can, y, ImportName, ImportCRUEI,
                JobId, MSGId, PermitNumber, TouchTime
            )
            y -= 20

        table.wrapOn(can, 0, 0)
        table.drawOn(can, 10, y - table_height)
        y -= (table_height + 15)

        return y

    # ─────────────────────────────────────────────────────────────────
    def _draw_invoice(self, can, y, invoice, ImportName, ImportCRUEI, JobId, MSGId, PermitNumber, TouchTime):
        (
            InvoiceNo, TermType, ImportPartyCode,
            TICurrency, TIExRate, TIAmount, TISAmount,
            FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
            ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
            OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
            CIFSUMAmount, GSTPercentage, GSTSUMAmount
        ) = invoice

        can.setFont('Times-Bold', 9)
        y -= 20
        can.drawString(10, y, "-" * 165)

        y -= 15
        can.drawString(10, y, "INVOICE NUMBER:")
        can.setFont('Times-Roman', 9)
        can.drawString(110, y, self._safe_val(InvoiceNo))

        y -= 15
        can.setFont('Times-Bold', 9)
        can.drawString(10, y, "INVOICE TERM:")
        can.setFont('Times-Roman', 9)
        can.drawString(110, y, self._safe_val(TermType))

        y -= 10

        table_data = [
            ['SNO', 'ITEM', 'CURRENCY', 'EXCHG. RATE', 'PERC.', 'AMOUNT', 'AMOUNT (SGD)'],
            ["1", "TOTAL INVOICE",
             self._safe_val(TICurrency), self._safe_val(TIExRate),
             "", self._safe_val(TIAmount), self._safe_val(TISAmount)],

            ["2", "FREIGHT CHARGES",
             self._safe_val(FCCurrency), self._safe_val(FCExRate),
             self._safe_val(FCCharge), self._safe_val(FCAmount), self._safe_val(FCSAmount)],

            ["3", "INSURANCE",
             self._safe_val(ICCurrency), self._safe_val(ICExRate),
             self._safe_val(ICCharge), self._safe_val(ICAmount), self._safe_val(ICSAmount)],

            ["4", "OTHER CHARGES",
             self._safe_val(OTCCurrency), self._safe_val(OTCExRate),
             self._safe_val(OTCCharge), self._safe_val(OTCAmount), self._safe_val(OTCSAmount)],

            ["5", "CUSTOMS VALUE",
             "", "", "", "", self._safe_val(CIFSUMAmount)],

            ["6", "GST",
             "", "", self._safe_val(GSTPercentage), "", self._safe_val(GSTSUMAmount)],
        ]

        col_widths = [30, 110, 75, 75, 55, 80, 95]

        table = Table(table_data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('ALIGN',          (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME',       (0, 0), (-1,  0), 'Times-Bold'),
            ('FONTNAME',       (0, 1), (-1, -1), 'Times-Roman'),
            ('FONTSIZE',       (0, 0), (-1, -1), 9),
            ('GRID',           (0, 0), (-1, -1), 0.5, colors.black),
            ('BOX',            (0, 0), (-1, -1), 1,   colors.black),
            ('VALIGN',         (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 0), (-1,  0), [colors.lightgrey]),
        ]))

        row_height   = 18
        table_height = row_height * len(table_data)

        # ── Page break if not enough space ──
        if y - table_height < 60:
            can.showPage()
            y = 820
            y = self._draw_header(
                can, y, ImportName, ImportCRUEI,
                JobId, MSGId, PermitNumber, TouchTime
            )
            y -= 20

        table.wrapOn(can, 0, 0)
        table.drawOn(can, 10, y - table_height)
        y -= (table_height + 15)

        return y

# # XML BUILD


# NS = {
#     "cbc": "urn:crimsonlogic:tn:schema:xsd:CommonBasicComponents-2",
#     "cac": "urn:crimsonlogic:tn:schema:xsd:CommonAggregateComponents-2",
#     "inp": "urn:crimsonlogic:tn:schema:xsd:InNonPayment",
#     "ipt": "urn:crimsonlogic:tn:schema:xsd:InPayment",
#     "out": "urn:crimsonlogic:tn:schema:xsd:OutwardDeclaration",
#     "tnp": "urn:crimsonlogic:tn:schema:xsd:TranshipmentMovement",
#     "coo": "urn:crimsonlogic:tn:schema:xsd:CertificateOfOrigin",
#     "":    "urn:crimsonlogic:tn:schema:xsd:TradenetDeclaration",
# }
# for prefix, uri in NS.items():
#     register_namespace(prefix, uri)


# def _q(prefix, tag):
#     """Build a Clark-notation qualified tag name, e.g. cbc:MessageReference."""
#     return f"{prefix}:{tag}" if prefix else tag


# class XmlSubmit(APIView):

#     # Maps MessageType to (namespace prefix, root element local name)
#     MESSAGE_TYPE_MAP = {
#         "INPDEC": ("inp", "InNonPayment"),
#         "IPTDEC": ("ipt", "InPayment"),
#         "OUTDEC": ("out", "Outward"),
#         "TNPDEC": ("tnp", "TranshipmentMovement"),
#         "COODEC": ("coo", "CertificateOfOrigin"),
#     }

#     # For refund/amend "update" submissions — different root element per type
#     UPDATE_MESSAGE_TYPE_MAP = {
#         "IPTDEC": ("ipt", "InPaymentUpdate"),
#         "INPDEC": ("inp", "InNonPaymentUpdate"),
#         "OUTDEC": ("out", "OutwardUpdate"),
#         "TNPDEC": ("tnp", "TranshipmentMovementUpdate"),
#     }

#     def get(self, request):
#         try:
#             permit_numbers = json.loads(request.GET.get("PermitNumber"))
#             username = request.GET.get("user")

#             if not permit_numbers:
#                 return JsonResponse({"error": "No permits selected"}, status=400)
#             if not username:
#                 return JsonResponse({"error": "User required"}, status=400)

#             TouchUser = username.upper()
#             now       = datetime.now()
#             TouchTime = now.strftime("%Y-%m-%d %H:%M:%S")

#             xml_results = []

#             with connection.cursor() as cursor:
#                 cursor.execute(
#                     "SELECT AccountId, MailBoxId FROM ManageUser WHERE UserName = %s",
#                     [TouchUser]
#                 )
#                 user_row = cursor.fetchone()
#             if not user_row:
#                 return JsonResponse({"error": f"User '{TouchUser}' not found"}, status=404)
#             AccountId, MailId = user_row

#             with connection.cursor() as cursor:
#                 cursor.execute(
#                     """SELECT CRUEI, DeclarantName, DeclarantCode, DeclarantTel, name, name1
#                        FROM DeclarantCompany WHERE TradeNetMailboxID = %s""",
#                     [MailId]
#                 )
#                 declarant_row = cursor.fetchone()
#             if not declarant_row:
#                 return JsonResponse({"error": f"DeclarantCompany not found for mailbox '{MailId}'"}, status=404)

#             DeclarantCompany = {
#                 "CRUEI": declarant_row[0],
#                 "DeclarantName": declarant_row[1],
#                 "DeclarantCode": declarant_row[2],
#                 "DeclarantTel": declarant_row[3],
#                 "Name": declarant_row[4],
#                 "Name1": declarant_row[5],
#             }

#             for permitNumber in permit_numbers:
#                 with connection.cursor() as cursor:
#                     cursor.execute("SELECT * FROM CommonHeaderTbl WHERE PermitId = %s", [permitNumber])
#                     columns  = [col[0] for col in cursor.description]
#                     HeadData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                 if not HeadData:
#                     return JsonResponse({"error": f"Permit {permitNumber} not found"}, status=404)

#                 head = HeadData[0]
#                 prmt_status   = (head.get("prmtStatus") or "").strip().upper()
#                 permit_status = (head.get("Status") or "").strip().upper()

#                 if prmt_status == "RFD" and permit_status == "NEW":
#                     # ── REFUND SUBMISSION ──────────────────────────────
#                     refund_rows = SqlDb.execute_query(
#                         "SELECT * FROM CommonRefund WHERE MSGId = %s", [head.get("MSGId")]
#                     )
#                     if not refund_rows:
#                         return JsonResponse(
#                             {"error": f"No refund data found for permit {permitNumber}"}, status=404
#                         )
#                     xml_data = self._build_refund_xml(permitNumber, HeadData, DeclarantCompany, refund_rows)

#                 elif prmt_status == "AMD" and permit_status == "NEW":
#                     # ── AMEND SUBMISSION ───────────────────────────────────────────
#                     amend_rows = SqlDb.execute_query(
#                         "SELECT * FROM CommonAmend WHERE MSGId = %s", [head.get("MSGId")]
#                     )
#                     if not amend_rows:
#                         return JsonResponse(
#                             {"error": f"No amend data found for permit {permitNumber}"}, status=404
#                         )

#                     with connection.cursor() as cursor:
#                         cursor.execute("SELECT * FROM CommonInvoiceDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         InvoiceData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                         cursor.execute("SELECT * FROM CommonItemDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         ItemData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                         cursor.execute("SELECT * FROM CommonContainerDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         ContainerData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                         cursor.execute("SELECT * FROM CommonCASCDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         CascData = [dict(zip(columns, row)) for row in cursor.fetchall()]


#                         cursor.execute(
#                             "SELECT * FROM CommonCPCDtl WHERE PermitId = %s ORDER BY RowNo",
#                             [permitNumber]
#                         )
#                         columns = [col[0] for col in cursor.description]
#                         CpcData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                     xml_data = self._build_amend_xml(
#                         permitNumber, TouchTime, AccountId, MailId, DeclarantCompany,
#                         HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData,
#                         amend_rows
#                     )

#                 else:
#                     # ── NORMAL DECLARATION SUBMISSION (unchanged) ──────
#                     with connection.cursor() as cursor:
#                         cursor.execute("SELECT * FROM CommonInvoiceDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         InvoiceData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                         cursor.execute("SELECT * FROM CommonItemDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         ItemData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                         cursor.execute("SELECT * FROM CommonContainerDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         ContainerData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                         cursor.execute("SELECT * FROM CommonCASCDtl WHERE PermitId = %s", [permitNumber])
#                         columns = [col[0] for col in cursor.description]
#                         CascData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                         cursor.execute(
#                             "SELECT * FROM CommonCPCDtl WHERE PermitId = %s ORDER BY RowNo",
#                             [permitNumber]
#                         )
#                         columns = [col[0] for col in cursor.description]
#                         CpcData = [dict(zip(columns, row)) for row in cursor.fetchall()]

#                     xml_data = self._build_xml(
#                         permitNumber, TouchTime, AccountId, MailId, DeclarantCompany,
#                         HeadData, InvoiceData, ItemData,
#                         ContainerData, CascData, CpcData
#                     )

#                 # ── SAVE XML TO D:\XML Output\<MailboxSuffix> ─────────
#                 mailbox_id = (head.get("TradeNetMailboxID") or "").strip()
#                 mailbox_suffix = mailbox_id.split(".")[-1].strip() if "." in mailbox_id else mailbox_id
#                 if not mailbox_suffix:
#                     mailbox_suffix = "UNKNOWN"

#                 msg_type = (head.get("MessageType") or "").strip()
#                 msg_id   = (head.get("MSGId") or "").strip()
#                 xml_filename = f"{msg_type}{msg_id}.xml" if (msg_type or msg_id) else f"{permitNumber}.xml"

#                 output_dir = os.path.join(r"D:\XML Output", mailbox_suffix)
#                 os.makedirs(output_dir, exist_ok=True)
#                 file_path = os.path.join(output_dir, xml_filename)
#                 with open(file_path, "wb") as f:
#                     f.write(xml_data)

#                 # ── UPDATE STATUS TO PEN on BOTH header tables ────────
#                 with connection.cursor() as cursor:
#                     cursor.execute("""
#                         UPDATE CommonHeaderTbl
#                         SET Status = 'PEN',
#                             TouchUser = %s, TouchTime = %s
#                         WHERE PermitId = %s
#                     """, [TouchUser, TouchTime, permitNumber])
#                     try:
#                         cursor.execute("""
#                             UPDATE InHeaderTbl
#                             SET Status = 'PEN',
#                                 TouchUser = %s, TouchTime = %s
#                             WHERE PermitId = %s
#                         """, [TouchUser, TouchTime, permitNumber])
#                     except Exception as err:
#                         print(f"Warning updating InHeaderTbl for {permitNumber}: {err}")

#                 xml_results.append({
#                     "permitNumber": permitNumber,
#                     "xmlFilename": xml_filename,
#                     "xml": xml_data.decode("utf-8")
#                 })

#             return JsonResponse({
#                 "SUCCESS": True,
#                 "message": f"{len(xml_results)} permit(s) submitted and saved to folder",
#                 "results": [r["xmlFilename"] for r in xml_results]
#             })
#         except Exception as e:
#             import traceback
#             traceback.print_exc()
#             return JsonResponse({"error": str(e)}, status=500)

# #     # ─────────────────────────────────────────────────────────────────
#     @staticmethod
#     def _txt(value):
#         """Safe string conversion for XML text — handles date/decimal/None."""
#         if value is None or value == "":
#             return ""
#         return str(value)

#     @staticmethod
#     def _set(el, value):
#         """Set element text; None/'' leaves it unset so tostring self-closes the tag."""
#         s = XmlSubmit._txt(value)
#         if s:
#             el.text = s

#     @staticmethod
#     def _code_only(value):
#         """'5 : Other non-Containerized' -> '5'. Leaves plain codes untouched."""
#         if value is None:
#             return None
#         value = str(value)
#         return value.split(":")[0].strip() if ":" in value else value.strip()

#     @staticmethod
#     def _yyyymmdd(value):
#         """Convert a date/datetime/str to YYYYMMDD, or '' if empty."""
#         if not value:
#             return ""
#         if isinstance(value, datetime):
#             return value.strftime("%Y%m%d")
#         s = str(value).strip()
#         for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
#             try:
#                 return datetime.strptime(s, fmt).strftime("%Y%m%d")
#             except ValueError:
#                 continue
#         # Last resort: strip separators, hope it's already close to YYYYMMDD
#         return s.replace("-", "").replace("/", "")[:8]
    
#     @staticmethod
#     def _cpc_carref(message_type, co_type):
#         """
#         Maps MessageType (+COType for OUT) to the CarRef column in CPCCodeValue.
#         IPT = InPayment, INP = InNonPayment, OUT w/o co / OUT with co = Outward
#         (with/without Certificate of Origin), TNP = Transhipment.
#         """
#         mt = (message_type or "").strip().upper()
#         if mt == "IPTDEC":
#             return "IPT"
#         if mt == "INPDEC":
#             return "INP"
#         if mt == "OUTDEC":
#             co = (co_type or "").strip()
#             return "OUT with co" if co not in ("", "--Select--", "None") else "OUT w/o co"
#         if mt == "TNPDEC":
#             return "TNP"
#         # COODEC (Certificate of Origin) has no CarRef mapping in the sample data
#         return ""

#     @staticmethod
#     def _cpc_pcode(cpc_type):
#         """
#         Normalizes CommonCPCDtl.CPCType to match CPCCodeValue.PCode.
#         Handles the one known spelling mismatch: 'SEASTORE' vs 'SEA STORE'
#         (CPCCodeValue uses both inconsistently — SEASTORE for INP/TNP rows,
#         'SEA STORE' for one specific OUT w/o co row). Default to the stored
#         value as-is; only special-case what's confirmed in the reference dump.
#         """
#         val = (cpc_type or "").strip().upper()
#         return val

#     @staticmethod
#     def _lookup_cpc_code(pcode, carref, dectype):
#         """
#         Looks up CPCCode from CPCCodeValue by (PCode, CarRef, DecType).
#         Returns None if no match — caller decides how to handle a miss.
#         """
#         if not pcode or not carref or not dectype:
#             return None
#         rows = SqlDb.execute_query(
#             """
#             SELECT CPCCode FROM CPCCodeValue
#             WHERE PCode = %s AND CarRef = %s AND DecType = %s
#             """,
#             [pcode, carref, dectype]
#         )
#         if rows:
#             return rows[0].get("CPCCode")

#         # Fallback: retry with the 'SEA STORE' spelling if the normal one missed
#         if pcode == "SEASTORE":
#             rows = SqlDb.execute_query(
#                 """
#                 SELECT CPCCode FROM CPCCodeValue
#                 WHERE PCode = %s AND CarRef = %s AND DecType = %s
#                 """,
#                 ["SEA STORE", carref, dectype]
#             )
#             if rows:
#                 return rows[0].get("CPCCode")

#         return None

#     @staticmethod
#     def _add(parent, tag, value):
#         """
#         Append a leaf element ONLY if `value` has meaningful (non-blank) text.
#         Returns the created element, or None if nothing was added.
#         """
#         text = XmlSubmit._txt(value)
#         if text is None or text.strip() == "":
#             return None
#         el = SubElement(parent, tag)
#         el.text = text
#         return el


#     @staticmethod
#     def _add_attr(parent, tag, value, **attrs):
#         """
#         Same as _add(), but also sets attributes (unitCode, currencyID, etc.)
#         on the element — and only if the element itself was actually created.
#         Blank/None attribute values are skipped too.
#         """
#         text = XmlSubmit._txt(value)
#         if text is None or text.strip() == "":
#             return None
#         el = SubElement(parent, tag)
#         el.text = text
#         for k, v in attrs.items():
#             if v is not None and str(v).strip() != "":
#                 el.set(k, str(v))
#         return el


#     @staticmethod
#     def _prune(el, parent):
#         """
#         Remove `el` from `parent` if it ended up with no child elements and
#         no text of its own (i.e. it would otherwise serialize as an empty
#         container tag like <cac:ReleaseLocation></cac:ReleaseLocation>).
#         Safe to call with el=None.
#         """
#         if el is None:
#             return
#         if len(el) == 0 and (el.text is None or el.text.strip() == ""):
#             try:
#                 parent.remove(el)
#             except ValueError:
#                 pass

#     # Supporting document
#     def _add_supporting_documents(self, parent, PermitId, file_type=None, fixed_doc_id=None):
#         query = "SELECT Name, DocumentType FROM CommonFile WHERE PermitId = %s"
#         params = [PermitId]
#         if file_type:
#             query += " AND Type = %s"
#             params.append(file_type)
#         query += " ORDER BY Sno"

#         files = SqlDb.execute_query(query, params)
#         if not files:
#             return

#         for f in files:
#             if fixed_doc_id:
#                 doc_id = fixed_doc_id
#             else:
#                 doc_type_raw = self._txt(f.get("DocumentType")).strip()
#                 doc_id = doc_type_raw.split(":")[0].strip() if ":" in doc_type_raw else doc_type_raw

#             file_name = self._txt(f.get("Name")).strip()
#             if not file_name:
#                 continue

#             doc_ref = SubElement(parent, _q("cac", "SupportingDocumentReference"))
#             self._add(doc_ref, _q("cbc", "DocumentID"), doc_id)
#             self._add(doc_ref, _q("cbc", "Filename"), file_name)
#             self._prune(doc_ref, parent)

#     # ---------------------------------------------------------------------------
#     # 2) REPLACEMENT for XmlSubmit._build_xml
#     # ---------------------------------------------------------------------------

#     def _build_xml(self, PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
#         HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData=None):
#         CpcData = CpcData or []

#         head = HeadData[0]
#         inv  = InvoiceData[0] if InvoiceData else {}
#         _t   = self._txt

#         msg_id = _t(head.get("MSGId")).strip()
#         urn_date = msg_id[:8] if len(msg_id) >= 12 else ""
#         urn_seq  = msg_id[8:] if len(msg_id) >= 12 else ""

#         # ── Determine namespace prefix + root element from MessageType ──
#         message_type = _t(head.get("MessageType")).strip().upper()
#         if message_type not in self.MESSAGE_TYPE_MAP:
#             raise ValueError(
#                 f"Unsupported MessageType '{message_type}' for PermitId {PermitId} — "
#                 f"no XML mapping defined. Supported types: {list(self.MESSAGE_TYPE_MAP.keys())}"
#             )
#         prefix, root_local = self.MESSAGE_TYPE_MAP[message_type]

#         root = Element(_q("", "TradenetDeclaration"))
#         root.set("xmlns:cbc", NS["cbc"])
#         root.set("xmlns:cac", NS["cac"])
#         root.set("xmlns:inp", NS["inp"])
#         root.set("xmlns:ipt", NS["ipt"])
#         root.set("xmlns:out", NS["out"])
#         root.set("xmlns:tnp", NS["tnp"])
#         root.set("xmlns:coo", NS["coo"])
#         root.set("xmlns", NS[""])

#         root.set("dateTime", datetime.now().strftime("%Y%m%d%H%M"))
#         root.set("instanceIdentifier", _t(head.get("JobId")))

#         self._add(root, _q("cbc", "MessageVersion"), "041")
#         self._add(root, _q("cbc", "SenderID"), head.get("TradeNetMailboxID"))
#         self._add(root, _q("cbc", "RecipientID"), "DCS4.DCS4001")
#         self._add(root, _q("cbc", "TotalNumberOfDeclaration"), "1")

#         inbound = SubElement(root, "InboundMessage")

#         # ─── MODULE ROOT (dynamic per message type) ──────────────────
#         module = SubElement(inbound, _q(prefix, root_local))

#         # ─── HEADER ─────────────────────────────────────────────────
#         hdr = SubElement(module, _q(prefix, "Header"))
#         self._add(hdr, _q("cbc", "MessageReference"), msg_id)

#         urn = SubElement(hdr, _q("cac", "UniqueReferenceNumber"))
#         self._add(urn, _q("cbc", "ID"), DeclarantCompany.get("CRUEI"))
#         self._add(urn, _q("cbc", "Date"), urn_date)
#         self._add(urn, _q("cbc", "SequenceNumeric"), urn_seq)
#         self._prune(urn, hdr)

#         self._add(hdr, _q("cbc", "DeclarantID"), head.get("TradeNetMailboxID"))
#         self._add(hdr, _q("cbc", "CommonAccessReference"), message_type)
#         self._add(hdr, _q("cbc", "DeclarationType"), self._code_only(head.get("DeclarationType")))
#         declare_ind = "true" if _t(head.get("DeclareIndicator")).upper() == "Y" else "false"
#         self._add(hdr, _q("cbc", "DeclarationIndicator"), declare_ind)

# #cpc datas 
#         if CpcData:
#             carref = self._cpc_carref(message_type, head.get("COType"))
#             dectype = self._code_only(head.get("DeclarationType"))

#             cpc = CpcData[0]  # only the first row is used — rows 2/3 ignored
#             pcode = self._cpc_pcode(cpc.get("CPCType"))

#             cpc_code = None
#             if pcode != "CNB":
#                 cpc_code = self._lookup_cpc_code(pcode, carref, dectype)

#             proc1 = _t(cpc.get("ProcessingCode1")).strip().upper()
#             proc2 = _t(cpc.get("ProcessingCode2")).strip().upper()
#             proc3 = _t(cpc.get("ProcessingCode3")).strip().upper()

#             cpc_info = SubElement(hdr, _q("cac", "CustomsProcedureCodeInformation"))
#             self._add(cpc_info, _q("cbc", "CustomsProcedureCode"), cpc_code)

#             proc_code = SubElement(cpc_info, _q("cac", "CPCProcessingCode"))
#             self._add(proc_code, _q("cbc", "ProcessingCodeOne"), proc1)
#             self._add(proc_code, _q("cbc", "ProcessingCodeTwo"), proc2)
#             self._add(proc_code, _q("cbc", "ProcessingCodeThree"), proc3)
#             self._prune(proc_code, cpc_info)

#             self._prune(cpc_info, hdr)


#         trade_remarks = _t(head.get("TradeRemarks")).strip()
#         if trade_remarks:
#             rme = SubElement(hdr, _q("cac", "Remarks"))
#             self._add(rme, _q("cbc", "FreeText"), trade_remarks)
#             self._prune(rme, hdr)

#         self._prune(hdr, module)

#         # ─── SUPPORTING DOCUMENTS ──────────────────────────────────
#         self._add_supporting_documents(module, PermitId)

#         # ─── CARGO ──────────────────────────────────────────────────
#         cargo = SubElement(module, _q(prefix, "Cargo"))
#         self._add(cargo, _q("cbc", "CargoPackingType"), self._code_only(head.get("CargoPackType")))

#         release_code = head.get("ReleaseLocation")
#         release = SubElement(cargo, _q("cac", "ReleaseLocation"))
#         self._add(release, _q("cbc", "LocationCode"), release_code)
#         self._add(release, _q("cbc", "LocationName"), head.get("ResLoaName"))
#         self._prune(release, cargo)

#         receipt_code = head.get("RecepitLocation")
#         receipt = SubElement(cargo, _q("cac", "ReceiptLocation"))
#         self._add(receipt, _q("cbc", "LocationCode"), receipt_code)
#         self._add(receipt, _q("cbc", "LocationName"), head.get("RecepitLocName"))
#         self._prune(receipt, cargo)

#       # ─── CONTAINERS (nested under Cargo, as TransportEquipment) ──
#         for cont in ContainerData:
#             equip = SubElement(cargo, _q("cac", "TransportEquipment"))
#             self._add(equip, _q("cbc", "SequenceNumeric"), cont.get("RowNo"))
#             container_no = _t(cont.get("ContainerNo")).strip().upper()
#             self._add(equip, _q("cbc", "EquipmentID"), container_no)
#             self._add(equip, _q("cbc", "SizeTypeCode"), self._code_only(cont.get("Size")))
#             self._add(equip, _q("cbc", "EquipmentWeightMeasureNumeric"), cont.get("Weight"))

#             seal_no = _t(cont.get("SealNo")).strip()
#             if seal_no:
#                 seal = SubElement(equip, _q("cac", "TransportEquipmentSeal"))
#                 self._add(seal, _q("cbc", "SealID"), seal_no)
#                 self._prune(seal, equip)

#             self._prune(equip, cargo)

#         self._prune(cargo, module)


#         # ─── TRANSPORT — emit Inward and/or Outward, independently ──
#         transport = SubElement(module, _q(prefix, "Transport"))

#         inward_mode_raw  = head.get("InwardTransportMode")
#         outward_mode_raw = head.get("OutwardTransportMode")
#         has_inward  = inward_mode_raw and str(inward_mode_raw).strip() not in ("",  "--Select--")
#         has_outward = outward_mode_raw and str(outward_mode_raw).strip() not in ("",  "--Select--")

#         if has_inward:
#             inward = SubElement(transport, _q("cac", "InwardTransport"))
#             means = SubElement(inward, _q("cac", "TransportMeans"))
#             mode = SubElement(means, _q("cac", "TransportMode"))

#             mode_code = self._code_only(inward_mode_raw)
#             self._add(mode, _q("cbc", "ModeCode"), mode_code)

#             if mode_code == "1":       # SEA
#                 conveyance_ref = head.get("VoyageNumber")
#                 mawb_obl       = head.get("OceanBillofLadingNo")
#                 transport_id   = head.get("VesselName")
#             elif mode_code == "4":     # AIR
#                 conveyance_ref = head.get("FlightNO")
#                 mawb_obl       = head.get("MasterAirwayBill")
#                 transport_id   = head.get("VesselName")
#             else:
#                 conveyance_ref = head.get("ConveyanceRefNo")
#                 mawb_obl       = head.get("HBL")
#                 transport_id   = head.get("TransportId")

#             self._add(mode, _q("cbc", "ConveyanceReferenceNumber"), conveyance_ref)
#             self._add(mode, _q("cbc", "TransportIdentifier"), transport_id)
#             self._prune(mode, means)
#             self._add(means, _q("cbc", "MAWBOUCROBLNumber"), mawb_obl)
#             self._prune(means, inward)

#             self._add(inward, _q("cbc", "ArrivalDate"), self._yyyymmdd(head.get("ArrivalDate")))
#             self._add(inward, _q("cbc", "LoadingPort"), head.get("LoadingPortCode"))
#             self._prune(inward, transport)

#         if has_outward:
#             outward = SubElement(transport, _q("cac", "OutwardTransport"))
#             means = SubElement(outward, _q("cac", "TransportMeans"))
#             mode = SubElement(means, _q("cac", "TransportMode"))

#             mode_code = self._code_only(outward_mode_raw)
#             self._add(mode, _q("cbc", "ModeCode"), mode_code)

#             if mode_code == "1":       # SEA
#                 conveyance_ref = head.get("OutVoyageNumber")
#                 mawb_obl       = head.get("OutOceanBillofLadingNo")
#             elif mode_code == "4":     # AIR
#                 conveyance_ref = head.get("OutFlightNO")
#                 mawb_obl       = head.get("OutMasterAirwayBill")
#             else:
#                 conveyance_ref = head.get("OutConveyanceRefNo")
#                 mawb_obl       = head.get("outHAWB")

#             self._add(mode, _q("cbc", "ConveyanceReferenceNumber"), conveyance_ref)
#             self._prune(mode, means)
#             self._add(means, _q("cbc", "MAWBOUCROBLNumber"), mawb_obl)
#             self._prune(means, outward)

#             self._add(outward, _q("cbc", "DepartureDate"), self._yyyymmdd(head.get("DepartureDate")))
#             self._add(outward, _q("cbc", "DischargePort"), head.get("DischargePort"))
#             self._add(outward, _q("cbc", "FinalDestinationCountry"), head.get("FinalDestinationCountry"))
#             self._prune(outward, transport)

#         self._prune(transport, module)

#         # ─── PARTY ──────────────────────────────────────────────────
#         party = SubElement(module, _q(prefix, "Party"))

#         # DeclarantParty
#         declarant = SubElement(party, _q("cac", "DeclarantParty"))
#         person = SubElement(declarant, _q("cac", "PersonInformation"))
#         self._add(person, _q("cbc", "CodeValue"), head.get("DeclarantCompanyCode"))
#         self._add(person, _q("cbc", "Name"), DeclarantCompany.get("DeclarantName"))
#         self._prune(person, declarant)
#         self._add(declarant, _q("cbc", "Telephone"), DeclarantCompany.get("DeclarantTel"))
#         self._prune(declarant, party)

#         # DeclaringAgentParty
#         agent = SubElement(party, _q("cac", "DeclaringAgentParty"))
#         agent_id = SubElement(agent, _q("cac", "PartyIdentification"))
#         self._add(agent_id, _q("cbc", "ID"), DeclarantCompany.get("CRUEI"))
#         self._prune(agent_id, agent)
#         agent_name = SubElement(agent, _q("cac", "PartyName"))
#         agent_full_name = (_t(DeclarantCompany.get("Name")) + _t(DeclarantCompany.get("Name1"))).strip()
#         self._add(agent_name, _q("cbc", "Name"), agent_full_name)
#         self._prune(agent_name, agent)
#         self._prune(agent, party)

#         def _party_block(local_name, code, lookup_query, params, name_field, cruei_field):
#             """Generic helper for the repeated 'look up code -> ID/Name party block' pattern."""
#             if not code or str(code).strip() in ("", "--Select--"):
#                 return
#             rows = SqlDb.execute_query(lookup_query, params)
#             name_val  = rows[0].get(name_field, "") if rows else ""
#             cruei_val = rows[0].get(cruei_field, "") if rows else ""

#             block = SubElement(party, _q("cac", local_name))
#             pid = SubElement(block, _q("cac", "PartyIdentification"))
#             self._add(pid, _q("cbc", "ID"), cruei_val)
#             self._prune(pid, block)
#             pname = SubElement(block, _q("cac", "PartyName"))
#             self._add(pname, _q("cbc", "Name"), name_val)
#             self._prune(pname, block)
#             self._prune(block, party)

#         _party_block(
#             "ImporterParty", head.get("ImporterCompanyCode"),
#             "SELECT Name, Name1, CRUEI FROM CommonImporter WHERE Code = %s",
#             [head.get("ImporterCompanyCode")], "Name", "CRUEI"
#         )
#         _party_block(
#             "FreightForwarderParty", head.get("FreightForwarderCode"),
#             "SELECT Name, Name1, CRUEI FROM CommonFreightForwarder WHERE Code = %s",
#             [head.get("FreightForwarderCode")], "Name", "CRUEI"
#         )
#         _party_block(
#             "InwardCarrierAgentParty",
#             head.get("InwardCarrierAgentCode") or head.get("Inwardcarriercode"),
#             "SELECT Name, Name1, CRUEI FROM CommonInwardCarrierAgent WHERE Code = %s",
#             [head.get("InwardCarrierAgentCode") or head.get("Inwardcarriercode")], "Name", "CRUEI"
#         )
#         _party_block(
#             "OutwardCarrierAgentParty", head.get("OutwardCarrierAgentCode"),
#             "SELECT Name, Name1, CRUEI FROM CommonOutwardCarrierAgent WHERE Code = %s",
#             [head.get("OutwardCarrierAgentCode")], "Name", "CRUEI"
#         )
#         _party_block(
#             "ConsigneeParty", head.get("CONSIGNEECode"),
#             "SELECT ConsigneeName, ConsigneeName1, ConsigneeCRUEI FROM CommonConsignee WHERE ConsigneeCode = %s",
#             [head.get("CONSIGNEECode")], "ConsigneeName", "ConsigneeCRUEI"
#         )
#         _party_block(
#             "ClaimantParty", head.get("ClaimantPartyCode"),
#             "SELECT ClaimantName, ClaimantName1, CRUEI FROM CommonClaimantParty WHERE ClaimantCode = %s",
#             [head.get("ClaimantPartyCode")], "ClaimantName", "CRUEI"
#         )
#         _party_block(
#             "EndUserParty", head.get("EndUserCode"),
#             "SELECT EndUserName, EndUserName1, EndUserCRUEI FROM CommonEndUser WHERE EndUserCode = %s",
#             [head.get("EndUserCode")], "EndUserName", "EndUserCRUEI"
#         )
#         _party_block(
#             "ManufacturerParty", head.get("Manufacturer"),
#             "SELECT ManufacturerName, ManufacturerName1, ManufacturerCRUEI FROM CommonManufacturer WHERE ManufacturerCode = %s",
#             [head.get("Manufacturer")], "ManufacturerName", "ManufacturerCRUEI"
#         )
#         _party_block(
#             "ExporterParty", head.get("ExporterCompanyCode"),
#             "SELECT Name, Name1, CRUEI FROM CommonExporter WHERE Code = %s",
#             [head.get("ExporterCompanyCode")], "Name", "CRUEI"
#         )
#         _party_block(
#             "HandlingAgentParty", head.get("HandlingAgentCode"),
#             "SELECT Name, Name1, CRUEI FROM CommonHandingAgent WHERE Code = %s",
#             [head.get("HandlingAgentCode")], "Name", "CRUEI"
#         )

#         license_ref = _t(head.get("License")).strip()
#         if license_ref and license_ref not in ("--Select--", "None"):
#             licence_el = SubElement(party, _q("cac", "Licence"))
#             self._add(licence_el, _q("cbc", "ReferenceID"), license_ref)
#             self._prune(licence_el, party)

#         self._prune(party, module)

#         # ─── INVOICES ───────────────────────────────────────────────
#         if message_type in ("IPTDEC", "INPDEC"):
#             for inv in InvoiceData:
#                 invoice = SubElement(module, _q("cac", "Invoice"))
#                 self._add(invoice, _q("cbc", "InvoiceNumber"), _t(inv.get("InvoiceNo")) or "-")
#                 self._add(invoice, _q("cbc", "InvoiceDate"), self._yyyymmdd(inv.get("InvoiceDate")))

#                 supplier_code = inv.get("SupplierCode")
#                 supplier_name = ""
#                 supplier_cruei = ""
#                 if supplier_code and str(supplier_code).strip() not in ("", "--Select--"):
#                     supplier_row = SqlDb.execute_query(
#                         "SELECT Name, CRUEI FROM CommonSupplierManufacturerPart WHERE Code = %s",
#                         [supplier_code]
#                     )
#                     supplier_cruei = supplier_row[0].get("CRUEI", "") if supplier_row else ""
#                     supplier_name = supplier_row[0].get("Name", "") if supplier_row else ""

#                 supplier = SubElement(invoice, _q("cac", "SupplierManufacturerParty"))
#                 self._add(supplier, _q("cbc", "CodeValue"), supplier_cruei)
#                 self._add(supplier, _q("cbc", "Name"), supplier_name)
#                 self._prune(supplier, invoice)

#                 self._add(invoice, _q("cbc", "UnitPriceTermType"), self._code_only(inv.get("TermType")))

#                 # ── TotalInvoiceValue (TI*) ──
#                 total_inv = SubElement(invoice, _q("cac", "TotalInvoiceValue"))
#                 self._add_attr(
#                     total_inv, _q("cbc", "Amount"), inv.get("TIAmount"),
#                     currencyID=_t(inv.get("TICurrency")) or "USD"
#                 )
#                 self._add(total_inv, _q("cbc", "ExchangeRate"), inv.get("TIExRate"))
#                 self._prune(total_inv, invoice)

#                 # ── FreightCharge (FC*) ──
#                 fc_amount = inv.get("FCAmount")
#                 if fc_amount not in (None, "", 0, "0", "0.00"):
#                     freight = SubElement(invoice, _q("cac", "FreightCharge"))
#                     self._add_attr(
#                         freight, _q("cbc", "Amount"), fc_amount,
#                         currencyID=_t(inv.get("FCCurrency")) or "SGD"
#                     )
#                     self._add(freight, _q("cbc", "ExchangeRate"), inv.get("FCExRate"))
#                     fc_charge = inv.get("FCCharge")
#                     if fc_charge not in (None, "", 0, "0", "0.00"):
#                         self._add(freight, _q("cbc", "ChargePercent"), fc_charge)
#                     self._prune(freight, invoice)

#                 # ── InsuranceCharge (IC*) ──
#                 ic_amount = inv.get("ICAmount")
#                 if ic_amount not in (None, "", 0, "0", "0.00"):
#                     insurance = SubElement(invoice, _q("cac", "InsuranceCharge"))
#                     self._add_attr(
#                         insurance, _q("cbc", "Amount"), ic_amount,
#                         currencyID=_t(inv.get("ICCurrency")) or "SGD"
#                     )
#                     self._add(insurance, _q("cbc", "ExchangeRate"), inv.get("ICExRate"))
#                     ic_charge = inv.get("ICCharge")
#                     if ic_charge not in (None, "", 0, "0", "0.00"):
#                         self._add(insurance, _q("cbc", "ChargePercent"), ic_charge)
#                     self._prune(insurance, invoice)

#                 # ── OtherCharge (OTC*) ── (unverified element name — kept from original)
#                 otc_amount = inv.get("OTCAmount")
#                 if otc_amount not in (None, "", 0, "0", "0.00"):
#                     other = SubElement(invoice, _q("cac", "OtherCharge"))
#                     self._add_attr(
#                         other, _q("cbc", "Amount"), otc_amount,
#                         currencyID=_t(inv.get("OTCCurrency")) or "SGD"
#                     )
#                     self._add(other, _q("cbc", "ExchangeRate"), inv.get("OTCExRate"))
#                     otc_charge = inv.get("OTCCharge")
#                     if otc_charge not in (None, "", 0, "0", "0.00"):
#                         self._add(other, _q("cbc", "ChargePercent"), otc_charge)
#                     self._prune(other, invoice)

#                 self._prune(invoice, module)

#         # ─── ITEMS ──────────────────────────────────────────────────
#         casc_by_item = {}
#         for c in CascData:
#             casc_by_item.setdefault(c.get("ItemNo"), []).append(c)

#         for it in ItemData:
#             item_el = SubElement(module, _q(prefix, "Item"))
#             self._add(item_el, _q("cbc", "ItemSequenceNumeric"), it.get("ItemNo"))
#             self._add(item_el, _q("cbc", "ItemHarmonizedSystemCode"), it.get("HSCode"))
#             self._add(item_el, _q("cbc", "GoodsDescription"), it.get("Description"))

#             qty = SubElement(item_el, _q("cac", "ItemQuantity"))
#             self._add_attr(
#                 qty, _q("cbc", "HarmonizedSystemQuantity"), it.get("HSQty"),
#                 unitCode=it.get("HSUOM")
#             )
#             self._prune(qty, item_el)

#             self._add(item_el, _q("cbc", "OriginCountry"), it.get("Contry"))

#             txn = SubElement(item_el, _q("cac", "TransactionValue"))
#             self._add(txn, _q("cbc", "ItemCIFFOBValue"), it.get("CIFFOB"))
#             self._prune(txn, item_el)

#             for casc in casc_by_item.get(it.get("ItemNo"), []):
#                 casc_el = SubElement(item_el, _q("cac", "CASCProduct"))
#                 self._add(casc_el, _q("cbc", "CASCProductCode"), casc.get("ProductCode"))
#                 self._add_attr(
#                     casc_el, _q("cbc", "CASCProductQuantity"), casc.get("Quantity"),
#                     unitCode=casc.get("ProductUOM")
#                 )

#                 # ── Additional CASC identification codes ──
#                 casc_code1 = _t(casc.get("CascCode1")).strip()
#                 casc_code2 = _t(casc.get("CascCode2")).strip()
#                 casc_code3 = _t(casc.get("CascCode3")).strip()
#                 if casc_code1 or casc_code2 or casc_code3:
#                     add_casc = SubElement(casc_el, _q("cac", "AdditionalCASCIdentification"))
#                     self._add(add_casc, _q("cbc", "CASCCodeOne"), casc_code1)
#                     self._add(add_casc, _q("cbc", "CASCCodeTwo"), casc_code2)
#                     self._add(add_casc, _q("cbc", "CASCCodeThree"), casc_code3)
#                     self._prune(add_casc, casc_el)

#                 self._prune(casc_el, item_el)

#             self._add(item_el, _q("cbc", "BrandName"), it.get("Brand"))
#             dg = "true" if _t(it.get("DGIndicator")).upper() == "Y" else "false"
#             self._add(item_el, _q("cbc", "DangerousGoodsIndicator"), dg)

#             in_hawb = _t(it.get("InHAWBOBL")).strip()
#             if in_hawb:
#                 self._add(item_el, _q("cbc", "InHAWBHUCRHBLNumber"), in_hawb)

#             self._add(item_el, _q("cbc", "ItemInvoiceNumber"), it.get("InvoiceNo"))

#             if message_type in ("IPTDEC", "INPDEC"):
#                 tariff = SubElement(item_el, _q("cac", "Tariff"))
#                 gst = SubElement(tariff, _q("cac", "GoodsAndServicesTax"))
#                 self._add(gst, _q("cbc", "GoodsAndServicesTaxPercent"), it.get("GSTRate"))
#                 self._add(gst, _q("cbc", "GoodsAndServicesTaxAmount"), it.get("GSTAmount"))
#                 self._prune(gst, tariff)
#                 self._prune(tariff, item_el)

#             self._prune(item_el, module)

#         # ─── SUMMARY ────────────────────────────────────────────────
#         summary = SubElement(module, _q(prefix, "Summary"))
#         self._add(summary, _q("cbc", "NumberOfItems"), len(ItemData) or head.get("NumberOfItems") or 0)
#         self._add(summary, _q("cbc", "TotalCIFFOBValue"), head.get("TotalCIFFOBValue"))

#         self._add_attr(
#             summary, _q("cbc", "TotalOuterPack"), head.get("TotalOuterPack"),
#             unitCode=head.get("TotalOuterPackUOM")
#         )
#         self._add_attr(
#             summary, _q("cbc", "TotalGrossWeight"), head.get("TotalGrossWeight"),
#             unitCode=head.get("TotalGrossWeightUOM")
#         )

#         if message_type in ("INPDEC", "IPTDEC"):
#             total_tariff = SubElement(summary, _q("cac", "TotalTariff"))
#             self._add(total_tariff, _q("cbc", "TotalGoodsAndServicesTaxAmount"), head.get("TotalGSTTaxAmt"))
#             total_amt_pay = _t(head.get("TotalAmtPay")).strip()
#             if total_amt_pay:
#                 self._add(total_tariff, _q("cbc", "TotalAmountPayable"), total_amt_pay)
#             self._prune(total_tariff, summary)

#         self._prune(summary, module)

#         return b'<?xml version="1.0" encoding="UTF-8"?>\n' + \
#             tostring(root, encoding="unicode").encode("utf-8")

# # build amend xml
#     def _build_amend_xml(self, PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
#                           HeadData, InvoiceData, ItemData, ContainerData, CascData,
#                           CpcData, AmendData):
#         head = HeadData[0]
#         amend = AmendData[0] if AmendData else {}
#         _t = self._txt

#         message_type = _t(head.get("MessageType")).strip().upper()
#         if message_type not in self.UPDATE_MESSAGE_TYPE_MAP:
#             raise ValueError(
#                 f"Unsupported MessageType '{message_type}' for amend XML on PermitId "
#                 f"{PermitId}. Supported: {list(self.UPDATE_MESSAGE_TYPE_MAP.keys())}"
#             )
#         update_prefix, update_root_local = self.UPDATE_MESSAGE_TYPE_MAP[message_type]
#         decl_prefix, decl_root_local = self.MESSAGE_TYPE_MAP[message_type]

#         # 1) Build the normal, fully field-mapped declaration XML — reuses
#         #    all of your existing party/invoice/item/summary logic as-is.
#         normal_xml_bytes = self._build_xml(
#             PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
#             HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData
#         )

#         # 2) Parse it back so we can lift out the declaration module and
#         #    re-parent it under the Update wrapper.
#         root = fromstring(normal_xml_bytes)

#         inbound_clark_tag = f"{{{NS['']}}}InboundMessage"
#         inbound = root.find(inbound_clark_tag)
#         if inbound is None:
#             raise ValueError(
#                 f"Could not locate InboundMessage element while building amend "
#                 f"XML for PermitId {PermitId}"
#             )

#         decl_clark_tag = f"{{{NS[decl_prefix]}}}{decl_root_local}"
#         decl_module = inbound.find(decl_clark_tag)
#         if decl_module is None:
#             raise ValueError(
#                 f"Could not locate declaration module '{decl_root_local}' while "
#                 f"building amend XML for PermitId {PermitId}"
#             )
#         inbound.remove(decl_module)

#         # 3) ipt:InPaymentUpdate (or the equivalent for the message type)
#         update_module = SubElement(inbound, _q(update_prefix, update_root_local))

#         # 3a) cac:Update block
#         update_el = SubElement(update_module, _q("cac", "Update"))
#         self._add(update_el, _q("cbc", "UpdateIndicatorCode"),
#                   amend.get("UpdateIndicator") or "AME")
#         self._add(update_el, _q("cbc", "UpdateRequestNumber"),
#                   amend.get("AmendmentCount") or "1")
#         self._add(update_el, _q("cbc", "UpdatePermitNumber"),
#                   amend.get("Permitno") or head.get("PermitNumber"))

#         reason = _t(amend.get("DescriptionOfReason")).strip()
#         if reason:
#             amendment_el = SubElement(update_el, _q("cac", "Amendment"))
#             self._add(amendment_el, _q("cbc", "AmendmentReason"), reason)
#             self._prune(amendment_el, update_el)
#         self._prune(update_el, update_module)

#         # 3b) ipt:Declaration wraps the full declaration content we built
#         #     in step 1 — moved over node-for-node, unchanged.
#         declaration_el = SubElement(update_module, _q(update_prefix, "Declaration"))
#         for child in list(decl_module):
#             declaration_el.append(child)

#         # 3c) Patch CommonAccessReference: normal declarations use the plain
#         #     MessageType ("IPTDEC"), but amend/update submissions must use
#         #     the *UPD variant ("IPTUPD") inside ipt:Header — same convention
#         #     _build_refund_xml already applies for refunds.
#         header_clark_tag = f"{{{NS[decl_prefix]}}}Header"
#         header_el = declaration_el.find(header_clark_tag)
#         if header_el is not None:
#             car_clark_tag = f"{{{NS['cbc']}}}CommonAccessReference"
#             car_el = header_el.find(car_clark_tag)
#             if car_el is not None:
#                 car_el.text = f"{message_type[:3]}UPD"

#         return b'<?xml version="1.0" encoding="UTF-8"?>\n' + \
#             tostring(root, encoding="unicode").encode("utf-8")

# #refund xml 
#     def _build_refund_xml(self, PermitId, HeadData, DeclarantCompany, RefundData):
#         head = HeadData[0]
#         refund = RefundData[0] if RefundData else {}
#         _t = self._txt

#         msg_id   = _t(head.get("MSGId")).strip()
#         urn_date = msg_id[:8] if len(msg_id) >= 12 else ""
#         urn_seq  = msg_id[8:] if len(msg_id) >= 12 else ""

#         message_type = _t(head.get("MessageType")).strip().upper()
#         if message_type not in self.UPDATE_MESSAGE_TYPE_MAP:
#             raise ValueError(
#                 f"Unsupported MessageType '{message_type}' for refund XML on PermitId {PermitId}. "
#                 f"Supported: {list(self.UPDATE_MESSAGE_TYPE_MAP.keys())}"
#             )
#         prefix, root_local = self.UPDATE_MESSAGE_TYPE_MAP[message_type]

#         root = Element(_q("", "TradenetDeclaration"))
#         root.set("xmlns:cbc", NS["cbc"])
#         root.set("xmlns:cac", NS["cac"])
#         root.set("xmlns:inp", NS["inp"])
#         root.set("xmlns:ipt", NS["ipt"])
#         root.set("xmlns:out", NS["out"])
#         root.set("xmlns:tnp", NS["tnp"])
#         root.set("xmlns:coo", NS["coo"])
#         root.set("xmlns", NS[""])
#         root.set("dateTime", datetime.now().strftime("%Y%m%d%H%M"))
#         root.set("instanceIdentifier", _t(head.get("JobId")))

#         self._add(root, _q("cbc", "MessageVersion"), "041")
#         self._add(root, _q("cbc", "SenderID"), head.get("TradeNetMailboxID"))
#         self._add(root, _q("cbc", "RecipientID"), "DCS4.DCS4001")
#         self._add(root, _q("cbc", "TotalNumberOfDeclaration"), "1")

#         inbound = SubElement(root, "InboundMessage")
#         module = SubElement(inbound, _q(prefix, root_local))

#         # ── cac:Update ──────────────────────────────────────────────
#         type_of_refund_raw = _t(refund.get("TypeOfRefund")).strip()
#         update_code = self._code_only(type_of_refund_raw)  # e.g. "FRF:Full refund" -> "FRF"

#         update = SubElement(module, _q("cac", "Update"))
#         self._add(update, _q("cbc", "UpdateIndicatorCode"), update_code)
#         # NOTE: sequence of update requests for this permit — no counter column
#         # exists yet in CommonRefund, defaulting to "1". Wire up a real counter
#         # (e.g. count of prior CommonRefund/CommonAmend rows for this PermitId)
#         # if TradeNet expects it to increment across resubmissions.
#         self._add(update, _q("cbc", "UpdateRequestNumber"), "1")
#         self._add(update, _q("cbc", "UpdatePermitNumber"), head.get("PermitNumber"))

#         reason_raw  = _t(refund.get("ReasonForRefund")).strip()
#         reason_code = self._code_only(reason_raw)  # e.g. "RF10:Double payment..." -> "RF10"
#         reason_desc = _t(refund.get("DescriptionOfReason")).strip()

#         refund_reason_el = SubElement(update, _q("cac", "Refund"))
#         self._add(refund_reason_el, _q("cbc", "ReasonCode"), reason_code)
#         self._add(refund_reason_el, _q("cbc", "Reason"), reason_desc)
#         self._prune(refund_reason_el, update)
#         self._prune(update, module)

#         # ── cac:RefundOnly ──────────────────────────────────────────
#         refund_only = SubElement(module, _q("cac", "RefundOnly"))

#         rhdr = SubElement(refund_only, _q("cac", "RefundHeader"))
#         self._add(rhdr, _q("cbc", "MessageReference"), msg_id)

#         urn = SubElement(rhdr, _q("cac", "UniqueReferenceNumber"))
#         self._add(urn, _q("cbc", "ID"), DeclarantCompany.get("CRUEI"))
#         self._add(urn, _q("cbc", "Date"), urn_date)
#         self._add(urn, _q("cbc", "SequenceNumeric"), urn_seq)
#         self._prune(urn, rhdr)

#         self._add(rhdr, _q("cbc", "DeclarantID"), head.get("TradeNetMailboxID"))
#         # "IPTUPD" for IPTDEC, "INPUPD" for INPDEC, etc.
#         self._add(rhdr, _q("cbc", "CommonAccessReference"), f"{message_type[:3]}UPD")
#         declare_ind = "true" if _t(head.get("DeclareIndicator")).upper() == "Y" else "false"
#         self._add(rhdr, _q("cbc", "DeclarationIndicator"), declare_ind)
#         self._prune(rhdr, refund_only)

#         declarant = SubElement(refund_only, _q("cac", "DeclarantParty"))
#         person = SubElement(declarant, _q("cac", "PersonInformation"))
#         self._add(person, _q("cbc", "CodeValue"), head.get("DeclarantCompanyCode"))
#         self._add(person, _q("cbc", "Name"), DeclarantCompany.get("DeclarantName"))
#         self._prune(person, declarant)
#         self._add(declarant, _q("cbc", "Telephone"), DeclarantCompany.get("DeclarantTel"))
#         self._prune(declarant, refund_only)

#         # Refund attachments — Type='RFD' files, fixed DocumentID "999"
#         self._add_supporting_documents(refund_only, PermitId, file_type="RFD", fixed_doc_id="999")

#         summary = SubElement(refund_only, _q("cac", "RefundSummary"))
#         total_tariff = SubElement(summary, _q("cac", "TotalTariffRefund"))
#         self._add(total_tariff, _q("cbc", "TotalGoodsAndServicesTaxRefundAmount"), refund.get("TotalGstAmt"))
#         self._prune(total_tariff, summary)
#         self._prune(summary, refund_only)

#         self._prune(refund_only, module)

#         return b'<?xml version="1.0" encoding="UTF-8"?>\n' + \
#             tostring(root, encoding="unicode").encode("utf-8")


NS = {
    "cbc": "urn:crimsonlogic:tn:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:crimsonlogic:tn:schema:xsd:CommonAggregateComponents-2",
    "inp": "urn:crimsonlogic:tn:schema:xsd:InNonPayment",
    "ipt": "urn:crimsonlogic:tn:schema:xsd:InPayment",
    "out": "urn:crimsonlogic:tn:schema:xsd:OutwardDeclaration",
    "tnp": "urn:crimsonlogic:tn:schema:xsd:TranshipmentMovement",
    "coo": "urn:crimsonlogic:tn:schema:xsd:CertificateOfOrigin",
    "":    "urn:crimsonlogic:tn:schema:xsd:TradenetDeclaration",
}
for prefix, uri in NS.items():
    register_namespace(prefix, uri)


def _q(prefix, tag):
    """Build a Clark-notation qualified tag name, e.g. cbc:MessageReference."""
    return f"{prefix}:{tag}" if prefix else tag


class XmlSubmit(APIView):

    # Maps MessageType to (namespace prefix, root element local name)
    MESSAGE_TYPE_MAP = {
        "INPDEC": ("inp", "InNonPayment"),
        "IPTDEC": ("ipt", "InPayment"),
        "OUTDEC": ("out", "Outward"),
        "TNPDEC": ("tnp", "TranshipmentMovement"),
        "COODEC": ("coo", "CertificateOfOrigin"),
    }

    # For refund/amend "update" submissions — different root element per type
    UPDATE_MESSAGE_TYPE_MAP = {
        "IPTDEC": ("ipt", "InPaymentUpdate"),
        "INPDEC": ("inp", "InNonPaymentUpdate"),
        "OUTDEC": ("out", "OutwardUpdate"),
        "TNPDEC": ("tnp", "TranshipmentMovementUpdate"),
    }

    # ════════════════════════════════════════════════════════════════════
    #  ENTRY POINT
    # ════════════════════════════════════════════════════════════════════
    def get(self, request):
        try:
            permit_numbers = json.loads(request.GET.get("PermitNumber"))
            username = request.GET.get("user")

            if not permit_numbers:
                return JsonResponse({"error": "No permits selected"}, status=400)
            if not username:
                return JsonResponse({"error": "User required"}, status=400)

            TouchUser = username.upper()
            now       = datetime.now()
            TouchTime = now.strftime("%Y-%m-%d %H:%M:%S")

            xml_results = []

            with connection.cursor() as cursor:
                cursor.execute(
                    "SELECT AccountId, MailBoxId FROM ManageUser WHERE UserName = %s",
                    [TouchUser]
                )
                user_row = cursor.fetchone()
            if not user_row:
                return JsonResponse({"error": f"User '{TouchUser}' not found"}, status=404)
            AccountId, MailId = user_row

            with connection.cursor() as cursor:
                cursor.execute(
                    """SELECT CRUEI, DeclarantName, DeclarantCode, DeclarantTel, name, name1
                       FROM DeclarantCompany WHERE TradeNetMailboxID = %s""",
                    [MailId]
                )
                declarant_row = cursor.fetchone()
            if not declarant_row:
                return JsonResponse({"error": f"DeclarantCompany not found for mailbox '{MailId}'"}, status=404)

            DeclarantCompany = {
                "CRUEI": declarant_row[0],
                "DeclarantName": declarant_row[1],
                "DeclarantCode": declarant_row[2],
                "DeclarantTel": declarant_row[3],
                "Name": declarant_row[4],
                "Name1": declarant_row[5],
            }

            for permitNumber in permit_numbers:
                with connection.cursor() as cursor:
                    cursor.execute("SELECT * FROM CommonHeaderTbl WHERE PermitId = %s", [permitNumber])
                    columns  = [col[0] for col in cursor.description]
                    HeadData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                if not HeadData:
                    return JsonResponse({"error": f"Permit {permitNumber} not found"}, status=404)

                head = HeadData[0]
                prmt_status   = (head.get("prmtStatus") or "").strip().upper()
                permit_status = (head.get("Status") or "").strip().upper()

                if prmt_status == "RFD" and permit_status == "NEW":
                    # ── REFUND SUBMISSION ──────────────────────────────
                    refund_rows = SqlDb.execute_query(
                        "SELECT * FROM CommonRefund WHERE MSGId = %s", [head.get("MSGId")]
                    )
                    if not refund_rows:
                        return JsonResponse(
                            {"error": f"No refund data found for permit {permitNumber}"}, status=404
                        )
                    xml_data = self._build_refund_xml(permitNumber, HeadData, DeclarantCompany, refund_rows)

                elif prmt_status == "CNL" and permit_status == "NEW":
                    # ── CANCEL SUBMISSION ───────────────────────────────
                    cancel_rows = SqlDb.execute_query(
                        "SELECT * FROM CommonCancel WHERE MSGId = %s", [head.get("MSGId")]
                    )
                    if not cancel_rows:
                        return JsonResponse(
                            {"error": f"No cancel data found for permit {permitNumber}"}, status=404
                        )
                    xml_data = self._build_cancel_xml(permitNumber, HeadData, DeclarantCompany, cancel_rows)

                elif prmt_status == "AMD" and permit_status == "NEW":
                    # ── AMEND SUBMISSION ───────────────────────────────
                    amend_rows = SqlDb.execute_query(
                        "SELECT * FROM CommonAmend WHERE MSGId = %s", [head.get("MSGId")]
                    )
                    if not amend_rows:
                        return JsonResponse(
                            {"error": f"No amend data found for permit {permitNumber}"}, status=404
                        )

                    with connection.cursor() as cursor:
                        cursor.execute("SELECT * FROM CommonInvoiceDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        InvoiceData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute("SELECT * FROM CommonItemDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        ItemData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute("SELECT * FROM CommonContainerDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        ContainerData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute("SELECT * FROM CommonCASCDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        CascData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute(
                            "SELECT * FROM CommonCPCDtl WHERE PermitId = %s ORDER BY RowNo",
                            [permitNumber]
                        )
                        columns = [col[0] for col in cursor.description]
                        CpcData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                    xml_data = self._build_amend_xml(
                        permitNumber, TouchTime, AccountId, MailId, DeclarantCompany,
                        HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData,
                        amend_rows
                    )

                else:
                    # ── NORMAL DECLARATION SUBMISSION ──────────────────
                    with connection.cursor() as cursor:
                        cursor.execute("SELECT * FROM CommonInvoiceDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        InvoiceData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute("SELECT * FROM CommonItemDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        ItemData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute("SELECT * FROM CommonContainerDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        ContainerData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute("SELECT * FROM CommonCASCDtl WHERE PermitId = %s", [permitNumber])
                        columns = [col[0] for col in cursor.description]
                        CascData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                        cursor.execute(
                            "SELECT * FROM CommonCPCDtl WHERE PermitId = %s ORDER BY RowNo",
                            [permitNumber]
                        )
                        columns = [col[0] for col in cursor.description]
                        CpcData = [dict(zip(columns, row)) for row in cursor.fetchall()]

                    xml_data = self._build_xml(
                        permitNumber, TouchTime, AccountId, MailId, DeclarantCompany,
                        HeadData, InvoiceData, ItemData,
                        ContainerData, CascData, CpcData
                    )

                # ── SAVE XML TO D:\XML Output\<MailboxSuffix> ─────────
                mailbox_id = (head.get("TradeNetMailboxID") or "").strip()
                mailbox_suffix = mailbox_id.split(".")[-1].strip() if "." in mailbox_id else mailbox_id
                if not mailbox_suffix:
                    mailbox_suffix = "UNKNOWN"

                msg_type = (head.get("MessageType") or "").strip()
                msg_id   = (head.get("MSGId") or "").strip()
                xml_filename = f"{msg_type}{msg_id}.xml" if (msg_type or msg_id) else f"{permitNumber}.xml"

                output_dir = os.path.join(r"D:\XML Output", mailbox_suffix)
                os.makedirs(output_dir, exist_ok=True)
                file_path = os.path.join(output_dir, xml_filename)
                with open(file_path, "wb") as f:
                    f.write(xml_data)

                # ── UPDATE STATUS TO PEN on BOTH header tables ────────
                with connection.cursor() as cursor:
                    cursor.execute("""
                        UPDATE CommonHeaderTbl
                        SET Status = 'PEN',
                            TouchUser = %s, TouchTime = %s
                        WHERE PermitId = %s
                    """, [TouchUser, TouchTime, permitNumber])
                    try:
                        cursor.execute("""
                            UPDATE InHeaderTbl
                            SET Status = 'PEN',
                                TouchUser = %s, TouchTime = %s
                            WHERE PermitId = %s
                        """, [TouchUser, TouchTime, permitNumber])
                    except Exception as err:
                        print(f"Warning updating InHeaderTbl for {permitNumber}: {err}")

                xml_results.append({
                    "permitNumber": permitNumber,
                    "xmlFilename": xml_filename,
                    "xml": xml_data.decode("utf-8")
                })

            return JsonResponse({
                "SUCCESS": True,
                "message": f"{len(xml_results)} permit(s) submitted and saved to folder",
                "results": [r["xmlFilename"] for r in xml_results]
            })
        except Exception as e:
            import traceback
            traceback.print_exc()
            return JsonResponse({"error": str(e)}, status=500)

    # ════════════════════════════════════════════════════════════════════
    #  SMALL TEXT/VALUE HELPERS
    # ════════════════════════════════════════════════════════════════════
    @staticmethod
    def _txt(value):
        """Safe string conversion for XML text — handles date/decimal/None."""
        if value is None or value == "":
            return ""
        return str(value)

    @staticmethod
    def _set(el, value):
        """Set element text; None/'' leaves it unset so tostring self-closes the tag."""
        s = XmlSubmit._txt(value)
        if s:
            el.text = s

    @staticmethod
    def _code_only(value):
        """'5 : Other non-Containerized' -> '5'. Leaves plain codes untouched."""
        if value is None:
            return None
        value = str(value)
        return value.split(":")[0].strip() if ":" in value else value.strip()

    @staticmethod
    def _yyyymmdd(value):
        """Convert a date/datetime/str to YYYYMMDD, or '' if empty."""
        if not value:
            return ""
        if isinstance(value, datetime):
            return value.strftime("%Y%m%d")
        s = str(value).strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(s, fmt).strftime("%Y%m%d")
            except ValueError:
                continue
        return s.replace("-", "").replace("/", "")[:8]

    @staticmethod
    def _cpc_carref(message_type, co_type):
        """
        Maps MessageType (+COType for OUT) to the CarRef column in CPCCodeValue.
        IPT = InPayment, INP = InNonPayment, OUT w/o co / OUT with co = Outward
        (with/without Certificate of Origin), TNP = Transhipment.
        """
        mt = (message_type or "").strip().upper()
        if mt == "IPTDEC":
            return "IPT"
        if mt == "INPDEC":
            return "INP"
        if mt == "OUTDEC":
            co = (co_type or "").strip()
            return "OUT with co" if co not in ("", "--Select--", "None") else "OUT w/o co"
        if mt == "TNPDEC":
            return "TNP"
        return ""

    @staticmethod
    def _cpc_pcode(cpc_type):
        """
        Normalizes CommonCPCDtl.CPCType to match CPCCodeValue.PCode.
        Default to the stored value as-is; only special-case what's confirmed
        in the reference dump (SEASTORE vs 'SEA STORE' handled in the lookup).
        """
        val = (cpc_type or "").strip().upper()
        return val

    @staticmethod
    def _lookup_cpc_code(pcode, carref, dectype):
        """
        Looks up CPCCode from CPCCodeValue by (PCode, CarRef, DecType).
        Returns None if no match — caller decides how to handle a miss.
        """
        if not pcode or not carref or not dectype:
            return None
        rows = SqlDb.execute_query(
            """
            SELECT CPCCode FROM CPCCodeValue
            WHERE PCode = %s AND CarRef = %s AND DecType = %s
            """,
            [pcode, carref, dectype]
        )
        if rows:
            return rows[0].get("CPCCode")

        if pcode == "SEASTORE":
            rows = SqlDb.execute_query(
                """
                SELECT CPCCode FROM CPCCodeValue
                WHERE PCode = %s AND CarRef = %s AND DecType = %s
                """,
                ["SEA STORE", carref, dectype]
            )
            if rows:
                return rows[0].get("CPCCode")

        return None

    # ════════════════════════════════════════════════════════════════════
    #  ELEMENT-BUILDING HELPERS (skip empty values, prune empty containers)
    # ════════════════════════════════════════════════════════════════════
    @staticmethod
    def _add(parent, tag, value):
        """
        Append a leaf element ONLY if `value` has meaningful (non-blank) text.
        Returns the created element, or None if nothing was added.
        """
        text = XmlSubmit._txt(value)
        if text is None or text.strip() == "":
            return None
        el = SubElement(parent, tag)
        el.text = text
        return el

    @staticmethod
    def _add_attr(parent, tag, value, **attrs):
        """
        Same as _add(), but also sets attributes (unitCode, currencyID, etc.)
        on the element — and only if the element itself was actually created.
        Blank/None attribute values are skipped too.
        """
        text = XmlSubmit._txt(value)
        if text is None or text.strip() == "":
            return None
        el = SubElement(parent, tag)
        el.text = text
        for k, v in attrs.items():
            if v is not None and str(v).strip() != "":
                el.set(k, str(v))
        return el

    @staticmethod
    def _prune(el, parent):
        """
        Remove `el` from `parent` if it ended up with no child elements and
        no text of its own (i.e. it would otherwise serialize as an empty
        container tag like <cac:ReleaseLocation></cac:ReleaseLocation>).
        Safe to call with el=None.
        """
        if el is None:
            return
        if len(el) == 0 and (el.text is None or el.text.strip() == ""):
            try:
                parent.remove(el)
            except ValueError:
                pass

    def _add_supporting_documents(self, parent, PermitId, file_type=None, fixed_doc_id=None):
        query = "SELECT Name, DocumentType FROM CommonFile WHERE PermitId = %s"
        params = [PermitId]
        if file_type:
            query += " AND Type = %s"
            params.append(file_type)
        query += " ORDER BY Sno"

        files = SqlDb.execute_query(query, params)
        if not files:
            return

        for f in files:
            if fixed_doc_id:
                doc_id = fixed_doc_id
            else:
                doc_type_raw = self._txt(f.get("DocumentType")).strip()
                doc_id = doc_type_raw.split(":")[0].strip() if ":" in doc_type_raw else doc_type_raw

            file_name = self._txt(f.get("Name")).strip()
            if not file_name:
                continue

            doc_ref = SubElement(parent, _q("cac", "SupportingDocumentReference"))
            self._add(doc_ref, _q("cbc", "DocumentID"), doc_id)
            self._add(doc_ref, _q("cbc", "Filename"), file_name)
            self._prune(doc_ref, parent)

    # ════════════════════════════════════════════════════════════════════
    #  NAMESPACE HELPER — the single, only place xmlns:* gets set
    # ════════════════════════════════════════════════════════════════════
    @staticmethod
    def _apply_namespaces(root):
        """
        Stamps all 8 xmlns:* declarations onto `root`. Called exactly once,
        right before the final tostring() in whichever builder is producing
        output — never before, never twice, and never on a tree that has
        already been through tostring()->fromstring() (which would cause
        ElementTree to auto-generate a second, colliding set based on which
        namespaces are actually "used" by Clark-notation tags in the tree).
        """
        root.set("xmlns:cbc", NS["cbc"])
        root.set("xmlns:cac", NS["cac"])
        root.set("xmlns:inp", NS["inp"])
        root.set("xmlns:ipt", NS["ipt"])
        root.set("xmlns:out", NS["out"])
        root.set("xmlns:tnp", NS["tnp"])
        root.set("xmlns:coo", NS["coo"])
        root.set("xmlns", NS[""])

    # ════════════════════════════════════════════════════════════════════
    #  NORMAL DECLARATION — thin wrapper over the shared tree builder
    # ════════════════════════════════════════════════════════════════════
    def _build_xml(self, PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
        HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData=None):
        root, inbound, module, prefix, message_type = self._build_declaration_tree(
            PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
            HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData
        )
        self._apply_namespaces(root)
        return b'<?xml version="1.0" encoding="UTF-8"?>\n' + \
            tostring(root, encoding="unicode").encode("utf-8")

    # ════════════════════════════════════════════════════════════════════
    #  SHARED TREE BUILDER — builds the full declaration as LIVE Elements,
    #  never serialized here. Used directly by _build_xml and reused
    #  (without any serialize/reparse round trip) by _build_amend_xml.
    # ════════════════════════════════════════════════════════════════════
    def _build_declaration_tree(self, PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
        HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData=None):
        CpcData = CpcData or []

        head = HeadData[0]
        inv  = InvoiceData[0] if InvoiceData else {}
        _t   = self._txt

        msg_id = _t(head.get("MSGId")).strip()
        urn_date = msg_id[:8] if len(msg_id) >= 12 else ""
        urn_seq  = msg_id[8:] if len(msg_id) >= 12 else ""

        message_type = _t(head.get("MessageType")).strip().upper()
        if message_type not in self.MESSAGE_TYPE_MAP:
            raise ValueError(
                f"Unsupported MessageType '{message_type}' for PermitId {PermitId} — "
                f"no XML mapping defined. Supported types: {list(self.MESSAGE_TYPE_MAP.keys())}"
            )
        prefix, root_local = self.MESSAGE_TYPE_MAP[message_type]

        # ── ROOT (no xmlns:* set here — that happens once, in the caller) ──
        root = Element(_q("", "TradenetDeclaration"))
        root.set("dateTime", datetime.now().strftime("%Y%m%d%H%M"))
        root.set("instanceIdentifier", _t(head.get("JobId")))

        self._add(root, _q("cbc", "MessageVersion"), "041")
        self._add(root, _q("cbc", "SenderID"), head.get("TradeNetMailboxID"))
        self._add(root, _q("cbc", "RecipientID"), "DCS4.DCS4001")
        self._add(root, _q("cbc", "TotalNumberOfDeclaration"), "1")

        inbound = SubElement(root, "InboundMessage")

        # ─── MODULE ROOT (dynamic per message type) ──────────────────
        module = SubElement(inbound, _q(prefix, root_local))

        # ─── HEADER ─────────────────────────────────────────────────
        hdr = SubElement(module, _q(prefix, "Header"))
        self._add(hdr, _q("cbc", "MessageReference"), msg_id)

        urn = SubElement(hdr, _q("cac", "UniqueReferenceNumber"))
        self._add(urn, _q("cbc", "ID"), DeclarantCompany.get("CRUEI"))
        self._add(urn, _q("cbc", "Date"), urn_date)
        self._add(urn, _q("cbc", "SequenceNumeric"), urn_seq)
        self._prune(urn, hdr)

        self._add(hdr, _q("cbc", "DeclarantID"), head.get("TradeNetMailboxID"))
        self._add(hdr, _q("cbc", "CommonAccessReference"), message_type)
        self._add(hdr, _q("cbc", "DeclarationType"), self._code_only(head.get("DeclarationType")))
        declare_ind = "true" if _t(head.get("DeclareIndicator")).upper() == "Y" else "false"
        self._add(hdr, _q("cbc", "DeclarationIndicator"), declare_ind)

        # ── CPC DATA ──
        if CpcData:
            carref = self._cpc_carref(message_type, head.get("COType"))
            dectype = self._code_only(head.get("DeclarationType"))

            cpc = CpcData[0]  # only the first row is used — rows 2/3 ignored
            pcode = self._cpc_pcode(cpc.get("CPCType"))

            cpc_code = None
            if pcode != "CNB":
                cpc_code = self._lookup_cpc_code(pcode, carref, dectype)

            proc1 = _t(cpc.get("ProcessingCode1")).strip().upper()
            proc2 = _t(cpc.get("ProcessingCode2")).strip().upper()
            proc3 = _t(cpc.get("ProcessingCode3")).strip().upper()

            cpc_info = SubElement(hdr, _q("cac", "CustomsProcedureCodeInformation"))
            self._add(cpc_info, _q("cbc", "CustomsProcedureCode"), cpc_code)

            proc_code = SubElement(cpc_info, _q("cac", "CPCProcessingCode"))
            self._add(proc_code, _q("cbc", "ProcessingCodeOne"), proc1)
            self._add(proc_code, _q("cbc", "ProcessingCodeTwo"), proc2)
            self._add(proc_code, _q("cbc", "ProcessingCodeThree"), proc3)
            self._prune(proc_code, cpc_info)

            self._prune(cpc_info, hdr)

        trade_remarks = _t(head.get("TradeRemarks")).strip()
        if trade_remarks:
            rme = SubElement(hdr, _q("cac", "Remarks"))
            self._add(rme, _q("cbc", "FreeText"), trade_remarks)
            self._prune(rme, hdr)

        self._prune(hdr, module)

        # ─── SUPPORTING DOCUMENTS ──────────────────────────────────
        self._add_supporting_documents(module, PermitId)

        # ─── CARGO ──────────────────────────────────────────────────
        cargo = SubElement(module, _q(prefix, "Cargo"))
        self._add(cargo, _q("cbc", "CargoPackingType"), self._code_only(head.get("CargoPackType")))

        release_code = head.get("ReleaseLocation")
        release = SubElement(cargo, _q("cac", "ReleaseLocation"))
        self._add(release, _q("cbc", "LocationCode"), release_code)
        self._add(release, _q("cbc", "LocationName"), head.get("ResLoaName"))
        self._prune(release, cargo)

        receipt_code = head.get("RecepitLocation")
        receipt = SubElement(cargo, _q("cac", "ReceiptLocation"))
        self._add(receipt, _q("cbc", "LocationCode"), receipt_code)
        self._add(receipt, _q("cbc", "LocationName"), head.get("RecepitLocName"))
        self._prune(receipt, cargo)

        # ─── CONTAINERS (nested under Cargo, as TransportEquipment) ──
        for cont in ContainerData:
            equip = SubElement(cargo, _q("cac", "TransportEquipment"))
            self._add(equip, _q("cbc", "SequenceNumeric"), cont.get("RowNo"))
            container_no = _t(cont.get("ContainerNo")).strip().upper()
            self._add(equip, _q("cbc", "EquipmentID"), container_no)
            self._add(equip, _q("cbc", "SizeTypeCode"), self._code_only(cont.get("Size")))
            self._add(equip, _q("cbc", "EquipmentWeightMeasureNumeric"), cont.get("Weight"))

            seal_no = _t(cont.get("SealNo")).strip()
            if seal_no:
                seal = SubElement(equip, _q("cac", "TransportEquipmentSeal"))
                self._add(seal, _q("cbc", "SealID"), seal_no)
                self._prune(seal, equip)

            self._prune(equip, cargo)

        self._prune(cargo, module)

        # ─── TRANSPORT — emit Inward and/or Outward, independently ──
        transport = SubElement(module, _q(prefix, "Transport"))

        inward_mode_raw  = head.get("InwardTransportMode")
        outward_mode_raw = head.get("OutwardTransportMode")
        has_inward  = inward_mode_raw and str(inward_mode_raw).strip() not in ("",  "--Select--")
        has_outward = outward_mode_raw and str(outward_mode_raw).strip() not in ("",  "--Select--")

        if has_inward:
            inward = SubElement(transport, _q("cac", "InwardTransport"))
            means = SubElement(inward, _q("cac", "TransportMeans"))
            mode = SubElement(means, _q("cac", "TransportMode"))

            mode_code = self._code_only(inward_mode_raw)
            self._add(mode, _q("cbc", "ModeCode"), mode_code)

            if mode_code == "1":       # SEA
                conveyance_ref = head.get("VoyageNumber")
                mawb_obl       = head.get("OceanBillofLadingNo")
                transport_id   = head.get("VesselName")
            elif mode_code == "4":     # AIR
                conveyance_ref = head.get("FlightNO")
                mawb_obl       = head.get("MasterAirwayBill")
                transport_id   = head.get("VesselName")
            else:
                conveyance_ref = head.get("ConveyanceRefNo")
                mawb_obl       = head.get("HBL")
                transport_id   = head.get("TransportId")

            self._add(mode, _q("cbc", "ConveyanceReferenceNumber"), conveyance_ref)
            self._add(mode, _q("cbc", "TransportIdentifier"), transport_id)
            self._prune(mode, means)
            self._add(means, _q("cbc", "MAWBOUCROBLNumber"), mawb_obl)
            self._prune(means, inward)

            self._add(inward, _q("cbc", "ArrivalDate"), self._yyyymmdd(head.get("ArrivalDate")))
            self._add(inward, _q("cbc", "LoadingPort"), head.get("LoadingPortCode"))
            self._prune(inward, transport)

        if has_outward:
            outward = SubElement(transport, _q("cac", "OutwardTransport"))
            means = SubElement(outward, _q("cac", "TransportMeans"))
            mode = SubElement(means, _q("cac", "TransportMode"))

            mode_code = self._code_only(outward_mode_raw)
            self._add(mode, _q("cbc", "ModeCode"), mode_code)

            if mode_code == "1":       # SEA
                conveyance_ref = head.get("OutVoyageNumber")
                mawb_obl       = head.get("OutOceanBillofLadingNo")
            elif mode_code == "4":     # AIR
                conveyance_ref = head.get("OutFlightNO")
                mawb_obl       = head.get("OutMasterAirwayBill")
            else:
                conveyance_ref = head.get("OutConveyanceRefNo")
                mawb_obl       = head.get("outHAWB")

            self._add(mode, _q("cbc", "ConveyanceReferenceNumber"), conveyance_ref)
            self._prune(mode, means)
            self._add(means, _q("cbc", "MAWBOUCROBLNumber"), mawb_obl)
            self._prune(means, outward)

            self._add(outward, _q("cbc", "DepartureDate"), self._yyyymmdd(head.get("DepartureDate")))
            self._add(outward, _q("cbc", "DischargePort"), head.get("DischargePort"))
            self._add(outward, _q("cbc", "FinalDestinationCountry"), head.get("FinalDestinationCountry"))
            self._prune(outward, transport)

        self._prune(transport, module)

        # ─── PARTY ──────────────────────────────────────────────────
        party = SubElement(module, _q(prefix, "Party"))

        declarant = SubElement(party, _q("cac", "DeclarantParty"))
        person = SubElement(declarant, _q("cac", "PersonInformation"))
        self._add(person, _q("cbc", "CodeValue"), head.get("DeclarantCompanyCode"))
        self._add(person, _q("cbc", "Name"), DeclarantCompany.get("DeclarantName"))
        self._prune(person, declarant)
        self._add(declarant, _q("cbc", "Telephone"), DeclarantCompany.get("DeclarantTel"))
        self._prune(declarant, party)

        agent = SubElement(party, _q("cac", "DeclaringAgentParty"))
        agent_id = SubElement(agent, _q("cac", "PartyIdentification"))
        self._add(agent_id, _q("cbc", "ID"), DeclarantCompany.get("CRUEI"))
        self._prune(agent_id, agent)
        agent_name = SubElement(agent, _q("cac", "PartyName"))
        agent_full_name = (_t(DeclarantCompany.get("Name")) + _t(DeclarantCompany.get("Name1"))).strip()
        self._add(agent_name, _q("cbc", "Name"), agent_full_name)
        self._prune(agent_name, agent)
        self._prune(agent, party)

        def _party_block(local_name, code, lookup_query, params, name_field, cruei_field):
            """Generic helper for the repeated 'look up code -> ID/Name party block' pattern."""
            if not code or str(code).strip() in ("", "--Select--"):
                return
            rows = SqlDb.execute_query(lookup_query, params)
            name_val  = rows[0].get(name_field, "") if rows else ""
            cruei_val = rows[0].get(cruei_field, "") if rows else ""

            block = SubElement(party, _q("cac", local_name))
            pid = SubElement(block, _q("cac", "PartyIdentification"))
            self._add(pid, _q("cbc", "ID"), cruei_val)
            self._prune(pid, block)
            pname = SubElement(block, _q("cac", "PartyName"))
            self._add(pname, _q("cbc", "Name"), name_val)
            self._prune(pname, block)
            self._prune(block, party)

        _party_block(
            "FreightForwarderParty", head.get("FreightForwarderCode"),
            "SELECT Name, Name1, CRUEI FROM CommonFreightForwarder WHERE Code = %s",
            [head.get("FreightForwarderCode")], "Name", "CRUEI"
        )
        _party_block(
            "InwardCarrierAgentParty",
            head.get("InwardCarrierAgentCode") or head.get("Inwardcarriercode"),
            "SELECT Name, Name1, CRUEI FROM CommonInwardCarrierAgent WHERE Code = %s",
            [head.get("InwardCarrierAgentCode") or head.get("Inwardcarriercode")], "Name", "CRUEI"
        )
        _party_block(
            "OutwardCarrierAgentParty", head.get("OutwardCarrierAgentCode"),
            "SELECT Name, Name1, CRUEI FROM CommonOutwardCarrierAgent WHERE Code = %s",
            [head.get("OutwardCarrierAgentCode")], "Name", "CRUEI"
        )
        _party_block(
            "ImporterParty", head.get("ImporterCompanyCode"),
            "SELECT Name, Name1, CRUEI FROM CommonImporter WHERE Code = %s",
            [head.get("ImporterCompanyCode")], "Name", "CRUEI"
        )
        _party_block(
            "ConsigneeParty", head.get("CONSIGNEECode"),
            "SELECT ConsigneeName, ConsigneeName1, ConsigneeCRUEI FROM CommonConsignee WHERE ConsigneeCode = %s",
            [head.get("CONSIGNEECode")], "ConsigneeName", "ConsigneeCRUEI"
        )
        _party_block(
            "ClaimantParty", head.get("ClaimantPartyCode"),
            "SELECT ClaimantName, ClaimantName1, CRUEI FROM CommonClaimantParty WHERE ClaimantCode = %s",
            [head.get("ClaimantPartyCode")], "ClaimantName", "CRUEI"
        )
        _party_block(
            "EndUserParty", head.get("EndUserCode"),
            "SELECT EndUserName, EndUserName1, EndUserCRUEI FROM CommonEndUser WHERE EndUserCode = %s",
            [head.get("EndUserCode")], "EndUserName", "EndUserCRUEI"
        )
        _party_block(
            "ManufacturerParty", head.get("Manufacturer"),
            "SELECT ManufacturerName, ManufacturerName1, ManufacturerCRUEI FROM CommonManufacturer WHERE ManufacturerCode = %s",
            [head.get("Manufacturer")], "ManufacturerName", "ManufacturerCRUEI"
        )
        _party_block(
            "ExporterParty", head.get("ExporterCompanyCode"),
            "SELECT Name, Name1, CRUEI FROM CommonExporter WHERE Code = %s",
            [head.get("ExporterCompanyCode")], "Name", "CRUEI"
        )
        _party_block(
            "HandlingAgentParty", head.get("HandlingAgentCode"),
            "SELECT Name, Name1, CRUEI FROM CommonHandingAgent WHERE Code = %s",
            [head.get("HandlingAgentCode")], "Name", "CRUEI"
        )

        license_ref = _t(head.get("License")).strip()
        if license_ref and license_ref not in ("--Select--", "None"):
            licence_el = SubElement(party, _q("cac", "Licence"))
            self._add(licence_el, _q("cbc", "ReferenceID"), license_ref)
            self._prune(licence_el, party)

        self._prune(party, module)

        # ─── INVOICES ───────────────────────────────────────────────
        if message_type in ("IPTDEC", "INPDEC"):
            for inv in InvoiceData:
                invoice = SubElement(module, _q("cac", "Invoice"))
                self._add(invoice, _q("cbc", "InvoiceNumber"), _t(inv.get("InvoiceNo")) or "-")
                self._add(invoice, _q("cbc", "InvoiceDate"), self._yyyymmdd(inv.get("InvoiceDate")))

                supplier_code = inv.get("SupplierCode")
                supplier_name = ""
                supplier_name1 = ""
                supplier_code_value = ""
                if supplier_code and str(supplier_code).strip() not in ("", "--Select--"):
                    supplier_row = SqlDb.execute_query(
                        "SELECT Name, Name1, Code FROM CommonSupplierManufacturerPart WHERE Code = %s",
                        [supplier_code]
                    )
                    if supplier_row:
                        supplier_name  = supplier_row[0].get("Name", "") or ""
                        supplier_name1 = supplier_row[0].get("Name1", "") or ""
                        supplier_code_value = supplier_row[0].get("Code", "") or supplier_code

                supplier = SubElement(invoice, _q("cac", "SupplierManufacturerParty"))
                self._add(supplier, _q("cbc", "CodeValue"), supplier_code_value)
                self._add(supplier, _q("cbc", "Name"), supplier_name)
                self._add(supplier, _q("cbc", "Name"), supplier_name1)
                self._prune(supplier, invoice)

                self._add(invoice, _q("cbc", "UnitPriceTermType"), self._code_only(inv.get("TermType")))

                total_inv = SubElement(invoice, _q("cac", "TotalInvoiceValue"))
                self._add_attr(
                    total_inv, _q("cbc", "Amount"), inv.get("TIAmount"),
                    currencyID=_t(inv.get("TICurrency")) or "USD"
                )
                self._add(total_inv, _q("cbc", "ExchangeRate"), inv.get("TIExRate"))
                self._prune(total_inv, invoice)

                fc_amount = inv.get("FCAmount")
                if fc_amount not in (None, "", 0, "0", "0.00"):
                    freight = SubElement(invoice, _q("cac", "FreightCharge"))
                    self._add_attr(
                        freight, _q("cbc", "Amount"), fc_amount,
                        currencyID=_t(inv.get("FCCurrency")) or "SGD"
                    )
                    self._add(freight, _q("cbc", "ExchangeRate"), inv.get("FCExRate"))
                    fc_charge = inv.get("FCCharge")
                    if fc_charge not in (None, "", 0, "0", "0.00"):
                        self._add(freight, _q("cbc", "ChargePercent"), fc_charge)
                    self._prune(freight, invoice)

                ic_amount = inv.get("ICAmount")
                if ic_amount not in (None, "", 0, "0", "0.00"):
                    insurance = SubElement(invoice, _q("cac", "InsuranceCharge"))
                    self._add_attr(
                        insurance, _q("cbc", "Amount"), ic_amount,
                        currencyID=_t(inv.get("ICCurrency")) or "SGD"
                    )
                    self._add(insurance, _q("cbc", "ExchangeRate"), inv.get("ICExRate"))
                    ic_charge = inv.get("ICCharge")
                    if ic_charge not in (None, "", 0, "0", "0.00"):
                        self._add(insurance, _q("cbc", "ChargePercent"), ic_charge)
                    self._prune(insurance, invoice)

                otc_amount = inv.get("OTCAmount")
                if otc_amount not in (None, "", 0, "0", "0.00"):
                    other = SubElement(invoice, _q("cac", "OtherCharge"))
                    self._add_attr(
                        other, _q("cbc", "Amount"), otc_amount,
                        currencyID=_t(inv.get("OTCCurrency")) or "SGD"
                    )
                    self._add(other, _q("cbc", "ExchangeRate"), inv.get("OTCExRate"))
                    otc_charge = inv.get("OTCCharge")
                    if otc_charge not in (None, "", 0, "0", "0.00"):
                        self._add(other, _q("cbc", "ChargePercent"), otc_charge)
                    self._prune(other, invoice)

                self._prune(invoice, module)

        # ─── ITEMS ──────────────────────────────────────────────────
        casc_by_item = {}
        for c in CascData:
            casc_by_item.setdefault(c.get("ItemNo"), []).append(c)

        for it in ItemData:
            item_el = SubElement(module, _q(prefix, "Item"))
            self._add(item_el, _q("cbc", "ItemSequenceNumeric"), it.get("ItemNo"))
            self._add(item_el, _q("cbc", "ItemHarmonizedSystemCode"), it.get("HSCode"))
            self._add(item_el, _q("cbc", "GoodsDescription"), it.get("Description"))

            qty = SubElement(item_el, _q("cac", "ItemQuantity"))
            self._add_attr(
                qty, _q("cbc", "HarmonizedSystemQuantity"), it.get("HSQty"),
                unitCode=it.get("HSUOM")
            )
            self._prune(qty, item_el)

            self._add(item_el, _q("cbc", "OriginCountry"), it.get("Contry"))

            txn = SubElement(item_el, _q("cac", "TransactionValue"))
            self._add(txn, _q("cbc", "ItemCIFFOBValue"), it.get("CIFFOB"))
            self._prune(txn, item_el)

            for casc in casc_by_item.get(it.get("ItemNo"), []):
                casc_el = SubElement(item_el, _q("cac", "CASCProduct"))
                self._add(casc_el, _q("cbc", "CASCProductCode"), casc.get("ProductCode"))
                self._add_attr(
                    casc_el, _q("cbc", "CASCProductQuantity"), casc.get("Quantity"),
                    unitCode=casc.get("ProductUOM")
                )

                casc_code1 = _t(casc.get("CascCode1")).strip()
                casc_code2 = _t(casc.get("CascCode2")).strip()
                casc_code3 = _t(casc.get("CascCode3")).strip()
                if casc_code1 or casc_code2 or casc_code3:
                    add_casc = SubElement(casc_el, _q("cac", "AdditionalCASCIdentification"))
                    self._add(add_casc, _q("cbc", "CASCCodeOne"), casc_code1)
                    self._add(add_casc, _q("cbc", "CASCCodeTwo"), casc_code2)
                    self._add(add_casc, _q("cbc", "CASCCodeThree"), casc_code3)
                    self._prune(add_casc, casc_el)

                self._prune(casc_el, item_el)

            self._add(item_el, _q("cbc", "BrandName"), it.get("Brand"))
            dg = "true" if _t(it.get("DGIndicator")).upper() == "Y" else "false"
            self._add(item_el, _q("cbc", "DangerousGoodsIndicator"), dg)

            in_hawb = _t(it.get("InHAWBOBL")).strip()
            if in_hawb:
                self._add(item_el, _q("cbc", "InHAWBHUCRHBLNumber"), in_hawb)

            self._add(item_el, _q("cbc", "ItemInvoiceNumber"), it.get("InvoiceNo"))

            if message_type in ("IPTDEC", "INPDEC"):
                tariff = SubElement(item_el, _q("cac", "Tariff"))
                gst = SubElement(tariff, _q("cac", "GoodsAndServicesTax"))
                self._add(gst, _q("cbc", "GoodsAndServicesTaxPercent"), it.get("GSTRate"))
                self._add(gst, _q("cbc", "GoodsAndServicesTaxAmount"), it.get("GSTAmount"))
                self._prune(gst, tariff)
                self._prune(tariff, item_el)

            self._prune(item_el, module)

        # ─── SUMMARY ────────────────────────────────────────────────
        summary = SubElement(module, _q(prefix, "Summary"))
        self._add(summary, _q("cbc", "NumberOfItems"), len(ItemData) or head.get("NumberOfItems") or 0)
        self._add(summary, _q("cbc", "TotalCIFFOBValue"), head.get("TotalCIFFOBValue"))

        self._add_attr(
            summary, _q("cbc", "TotalOuterPack"), head.get("TotalOuterPack"),
            unitCode=head.get("TotalOuterPackUOM")
        )
        self._add_attr(
            summary, _q("cbc", "TotalGrossWeight"), head.get("TotalGrossWeight"),
            unitCode=head.get("TotalGrossWeightUOM")
        )

        if message_type in ("INPDEC", "IPTDEC"):
            total_tariff = SubElement(summary, _q("cac", "TotalTariff"))
            self._add(total_tariff, _q("cbc", "TotalGoodsAndServicesTaxAmount"), head.get("TotalGSTTaxAmt"))
            total_amt_pay = _t(head.get("TotalAmtPay")).strip()
            if total_amt_pay:
                self._add(total_tariff, _q("cbc", "TotalAmountPayable"), total_amt_pay)
            self._prune(total_tariff, summary)

        self._prune(summary, module)

        # Return the LIVE tree — no serialization here. _build_xml (or
        # _build_amend_xml, reusing these same Elements) decides when to
        # stamp namespaces and call tostring().
        return root, inbound, module, prefix, message_type

    # ════════════════════════════════════════════════════════════════════
    #  AMEND SUBMISSION — reuses the shared tree, no reparse
    # ════════════════════════════════════════════════════════════════════
    def _build_amend_xml(self, PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
                          HeadData, InvoiceData, ItemData, ContainerData, CascData,
                          CpcData, AmendData):
        head = HeadData[0]
        amend = AmendData[0] if AmendData else {}
        _t = self._txt

        message_type = _t(head.get("MessageType")).strip().upper()
        if message_type not in self.UPDATE_MESSAGE_TYPE_MAP:
            raise ValueError(
                f"Unsupported MessageType '{message_type}' for amend XML on PermitId "
                f"{PermitId}. Supported: {list(self.UPDATE_MESSAGE_TYPE_MAP.keys())}"
            )
        update_prefix, update_root_local = self.UPDATE_MESSAGE_TYPE_MAP[message_type]

        # Build the full declaration tree ONCE as live Elements — no
        # serialize-then-reparse round trip. This is what avoids both the
        # "missing namespace" bug (fromstring() never runs, so nothing
        # strips xmlns:* or causes tostring() to selectively re-emit only
        # "used" prefixes) and the "attribute redefined" bug (tostring() is
        # only called once, at the very end of this method).
        root, inbound, decl_module, decl_prefix, _ = self._build_declaration_tree(
            PermitId, TouchTime, AccountId, MailId, DeclarantCompany,
            HeadData, InvoiceData, ItemData, ContainerData, CascData, CpcData
        )

        # Detach the plain declaration module (e.g. ipt:InPayment) — its
        # children get moved into ipt:Declaration under the Update wrapper.
        inbound.remove(decl_module)

        # ── ipt:InPaymentUpdate (or equivalent for the message type) ──────
        update_module = SubElement(inbound, _q(update_prefix, update_root_local))

        # ── cac:Update block ───────────────────────────────────────────
        update_el = SubElement(update_module, _q("cac", "Update"))
        self._add(update_el, _q("cbc", "UpdateIndicatorCode"),
                  amend.get("UpdateIndicator") or "AME")
        self._add(update_el, _q("cbc", "UpdateRequestNumber"),
                  amend.get("AmendmentCount") or "1")
        self._add(update_el, _q("cbc", "UpdatePermitNumber"),
                  amend.get("Permitno") or head.get("PermitNumber"))

        reason = _t(amend.get("DescriptionOfReason")).strip()
        if reason:
            amendment_el = SubElement(update_el, _q("cac", "Amendment"))
            self._add(amendment_el, _q("cbc", "AmendmentReason"), reason)
            self._prune(amendment_el, update_el)
        self._prune(update_el, update_module)

        # ── ipt:Declaration wraps the full declaration content — moved
        #    node-for-node (same Element objects, not copies) ────────────
        declaration_el = SubElement(update_module, _q(update_prefix, "Declaration"))
        for child in list(decl_module):
            declaration_el.append(child)

        # ── Patch CommonAccessReference: normal declarations use the plain
        #    MessageType ("IPTDEC"), but amend/update submissions must use
        #    the *UPD variant ("IPTUPD") inside ipt:Header ────────────────
        header_tag = _q(decl_prefix, "Header")
        header_el = declaration_el.find(header_tag)
        if header_el is not None:
            car_tag = _q("cbc", "CommonAccessReference")
            car_el = header_el.find(car_tag)
            if car_el is not None:
                car_el.text = f"{message_type[:3]}UPD"

        # Single, authoritative namespace application.
        self._apply_namespaces(root)

        return b'<?xml version="1.0" encoding="UTF-8"?>\n' + \
            tostring(root, encoding="unicode").encode("utf-8")


    # cancel xml
    # ════════════════════════════════════════════════════════════════════
        #  CANCEL SUBMISSION — same pattern as refund: single-pass tree build
    # ════════════════════════════════════════════════════════════════════
    def _build_cancel_xml(self, PermitId, HeadData, DeclarantCompany, CancelData):
        head = HeadData[0]
        cancel = CancelData[0] if CancelData else {}
        _t = self._txt

        msg_id   = _t(head.get("MSGId")).strip()
        urn_date = msg_id[:8] if len(msg_id) >= 12 else ""
        urn_seq  = msg_id[8:] if len(msg_id) >= 12 else ""

        message_type = _t(head.get("MessageType")).strip().upper()
        if message_type not in self.UPDATE_MESSAGE_TYPE_MAP:
            raise ValueError(
                f"Unsupported MessageType '{message_type}' for cancel XML on PermitId {PermitId}. "
                f"Supported: {list(self.UPDATE_MESSAGE_TYPE_MAP.keys())}"
            )
        prefix, root_local = self.UPDATE_MESSAGE_TYPE_MAP[message_type]

        root = Element(_q("", "TradenetDeclaration"))
        root.set("dateTime", datetime.now().strftime("%Y%m%d%H%M"))
        root.set("instanceIdentifier", _t(head.get("JobId")))

        self._add(root, _q("cbc", "MessageVersion"), "041")
        self._add(root, _q("cbc", "SenderID"), head.get("TradeNetMailboxID"))
        self._add(root, _q("cbc", "RecipientID"), "DCS4.DCS4001")
        self._add(root, _q("cbc", "TotalNumberOfDeclaration"), "1")

        inbound = SubElement(root, "InboundMessage")
        module = SubElement(inbound, _q(prefix, root_local))

        # ── cac:Update ──────────────────────────────────────────────
        update = SubElement(module, _q("cac", "Update"))
        self._add(update, _q("cbc", "UpdateIndicatorCode"), "CNL")
        self._add(update, _q("cbc", "UpdateRequestNumber"), "1")
        self._add(update, _q("cbc", "UpdatePermitNumber"), head.get("PermitNumber"))
        self._prune(update, module)

        # ── cac:Cancellation ────────────────────────────────────────
        cancellation = SubElement(module, _q("cac", "Cancellation"))

        chdr = SubElement(cancellation, _q("cac", "CancellationHeader"))
        self._add(chdr, _q("cbc", "MessageReference"), msg_id)

        urn = SubElement(chdr, _q("cac", "UniqueReferenceNumber"))
        self._add(urn, _q("cbc", "ID"), DeclarantCompany.get("CRUEI"))
        self._add(urn, _q("cbc", "Date"), urn_date)
        self._add(urn, _q("cbc", "SequenceNumeric"), urn_seq)
        self._prune(urn, chdr)

        self._add(chdr, _q("cbc", "DeclarantID"), head.get("TradeNetMailboxID"))
        # "INPUPD" for INPDEC, "IPTUPD" for IPTDEC, etc. — same convention as refund
        self._add(chdr, _q("cbc", "CommonAccessReference"), f"{message_type[:3]}UPD")
        declare_ind = "true" if _t(head.get("DeclareIndicator")).upper() == "Y" else "false"
        self._add(chdr, _q("cbc", "DeclarationIndicator"), declare_ind)

        # ReasonForCancel is stored as "C04:Some description" (from CommonMaster
        # TypeId=75) — same "code:description" convention as refund's ReasonCode.
        reason_raw  = _t(cancel.get("ReasonForCancel")).strip()
        reason_code = self._code_only(reason_raw)
        self._add(chdr, _q("cbc", "CancellationReasonCode"), reason_code)
        self._prune(chdr, cancellation)

        declarant = SubElement(cancellation, _q("cac", "DeclarantParty"))
        person = SubElement(declarant, _q("cac", "PersonInformation"))
        self._add(person, _q("cbc", "CodeValue"), DeclarantCompany.get("DeclarantCode"))
        self._add(person, _q("cbc", "Name"), DeclarantCompany.get("DeclarantName"))
        self._prune(person, declarant)
        self._add(declarant, _q("cbc", "Telephone"), DeclarantCompany.get("DeclarantTel"))
        self._prune(declarant, cancellation)

        # Cancel attachments — mirrors refund's pattern: Type='CNL' files,
        # fixed DocumentID "999". Your Cancel.jsx already uploads with
        # Type: "CNL", so this picks those up automatically.
        self._add_supporting_documents(cancellation, PermitId, file_type="CNL", fixed_doc_id="999")

        self._prune(cancellation, module)

        self._apply_namespaces(root)

        return b'<?xml version="1.0" encoding="UTF-8"?>\n' + \
            tostring(root, encoding="unicode").encode("utf-8")

    # ════════════════════════════════════════════════════════════════════
    #  REFUND SUBMISSION — builds its own small tree in one pass, never
    #  reparses, so it was never at risk of either namespace bug.
    # ════════════════════════════════════════════════════════════════════
    def _build_refund_xml(self, PermitId, HeadData, DeclarantCompany, RefundData):
        head = HeadData[0]
        refund = RefundData[0] if RefundData else {}
        _t = self._txt

        msg_id   = _t(head.get("MSGId")).strip()
        urn_date = msg_id[:8] if len(msg_id) >= 12 else ""
        urn_seq  = msg_id[8:] if len(msg_id) >= 12 else ""

        message_type = _t(head.get("MessageType")).strip().upper()
        if message_type not in self.UPDATE_MESSAGE_TYPE_MAP:
            raise ValueError(
                f"Unsupported MessageType '{message_type}' for refund XML on PermitId {PermitId}. "
                f"Supported: {list(self.UPDATE_MESSAGE_TYPE_MAP.keys())}"
            )
        prefix, root_local = self.UPDATE_MESSAGE_TYPE_MAP[message_type]

        root = Element(_q("", "TradenetDeclaration"))
        root.set("dateTime", datetime.now().strftime("%Y%m%d%H%M"))
        root.set("instanceIdentifier", _t(head.get("JobId")))

        self._add(root, _q("cbc", "MessageVersion"), "041")
        self._add(root, _q("cbc", "SenderID"), head.get("TradeNetMailboxID"))
        self._add(root, _q("cbc", "RecipientID"), "DCS4.DCS4001")
        self._add(root, _q("cbc", "TotalNumberOfDeclaration"), "1")

        inbound = SubElement(root, "InboundMessage")
        module = SubElement(inbound, _q(prefix, root_local))

        # ── cac:Update ──────────────────────────────────────────────
        type_of_refund_raw = _t(refund.get("TypeOfRefund")).strip()
        update_code = self._code_only(type_of_refund_raw)

        update = SubElement(module, _q("cac", "Update"))
        self._add(update, _q("cbc", "UpdateIndicatorCode"), update_code)
        self._add(update, _q("cbc", "UpdateRequestNumber"), "1")
        self._add(update, _q("cbc", "UpdatePermitNumber"), head.get("PermitNumber"))

        reason_raw = _t(refund.get("ReasonForREfund")).strip()
        if ":" in reason_raw:
            reason_code, reason_desc = reason_raw.split(":", 1)
            reason_code = reason_code.strip()
            reason_desc = reason_desc.strip().upper()
        else:
            reason_code = reason_raw
            reason_desc = ""

        refund_reason_el = SubElement(update, _q("cac", "Refund"))
        self._add(refund_reason_el, _q("cbc", "ReasonCode"), reason_code)
        self._add(refund_reason_el, _q("cbc", "Reason"), reason_desc)
        self._prune(refund_reason_el, update)

        # ── cac:RefundOnly ──────────────────────────────────────────
        refund_only = SubElement(module, _q("cac", "RefundOnly"))

        rhdr = SubElement(refund_only, _q("cac", "RefundHeader"))
        self._add(rhdr, _q("cbc", "MessageReference"), msg_id)

        urn = SubElement(rhdr, _q("cac", "UniqueReferenceNumber"))
        self._add(urn, _q("cbc", "ID"), DeclarantCompany.get("CRUEI"))
        self._add(urn, _q("cbc", "Date"), urn_date)
        self._add(urn, _q("cbc", "SequenceNumeric"), urn_seq)
        self._prune(urn, rhdr)

        self._add(rhdr, _q("cbc", "DeclarantID"), head.get("TradeNetMailboxID"))
        self._add(rhdr, _q("cbc", "CommonAccessReference"), f"{message_type[:3]}UPD")
        declare_ind = "true" if _t(head.get("DeclareIndicator")).upper() == "Y" else "false"
        self._add(rhdr, _q("cbc", "DeclarationIndicator"), declare_ind)
        self._prune(rhdr, refund_only)

        declarant = SubElement(refund_only, _q("cac", "DeclarantParty"))
        person = SubElement(declarant, _q("cac", "PersonInformation"))
        self._add(person, _q("cbc", "CodeValue"), DeclarantCompany.get("DeclarantCode"))
        self._add(person, _q("cbc", "Name"), DeclarantCompany.get("DeclarantName"))
        self._prune(person, declarant)
        self._add(declarant, _q("cbc", "Telephone"), DeclarantCompany.get("DeclarantTel"))
        self._prune(declarant, refund_only)

        self._add_supporting_documents(refund_only, PermitId, file_type="RFD", fixed_doc_id="999")

        summary = SubElement(refund_only, _q("cac", "RefundSummary"))
        total_tariff = SubElement(summary, _q("cac", "TotalTariffRefund"))
        self._add(total_tariff, _q("cbc", "TotalGoodsAndServicesTaxRefundAmount"), refund.get("TotalGstAmt"))
        self._prune(total_tariff, summary)
        self._prune(summary, refund_only)

        self._prune(refund_only, module)

        self._apply_namespaces(root)

        return b'<?xml version="1.0" encoding="UTF-8"?>\n' + \
            tostring(root, encoding="unicode").encode("utf-8")
            

# Amend Details
# Get Amend Details by MSGId
class GetAmendByMsgId(APIView):
    table = "CommonAmend"

    def get(self, request):
        try:
            msg_id = request.GET.get("MSGId")
            if not msg_id:
                return Response({"error": "MSGId is required"}, status=status.HTTP_400_BAD_REQUEST)
            query = f"SELECT * FROM {self.table} WHERE MSGId = %s"
            rows = SqlDb.execute_query(query, [msg_id])
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

# class PostAmendTable(APIView):
#     table = "CommonAmend"
#     allowed_columns = {
#         "Permitno", "AmendmentCount", "UpdateIndicator",
#         "ReplacementPermitno", "DescriptionOfReason",
#         "PermitExtension", "ExtendImportPeriod",
#         "DeclarationIndigator", "AmendType",
#         "TouchUser", "TouchTime", "MSGId",
#     }

#     def post(self, request):
#         item = request.data
#         if not item or not isinstance(item, dict):
#             return Response({"error": "No valid data provided"}, status=400)

#         item.pop("Id", None)
#         msg_id = item.get("MSGId", "")
#         permit_id = item.get("Permitno", "")

#         if not msg_id:
#             return Response({"error": "MSGId is required"}, status=400)

#         columns = [k for k in item.keys() if k in self.allowed_columns]
#         if not columns:
#             return Response({"error": "No valid columns provided"}, status=400)

#         try:
#             existing = SqlDb.execute_query(
#                 f"SELECT Id FROM {self.table} WHERE MSGId = %s", [msg_id]
#             )
#             if existing:
#                 update_cols = [c for c in columns if c != "MSGId"]
#                 set_clause = ", ".join([f"{c} = %s" for c in update_cols])
#                 values = [item[c] for c in update_cols] + [msg_id]
#                 SqlDb.execute_query(
#                     f"UPDATE {self.table} SET {set_clause} WHERE MSGId = %s", values
#                 )
#                 action = "updated"
#             else:
#                 col_str = ", ".join(columns)
#                 ph_str = ", ".join(["%s"] * len(columns))
#                 values = [item[c] for c in columns]
#                 SqlDb.execute_query(
#                     f"INSERT INTO {self.table} ({col_str}) VALUES ({ph_str})", values
#                 )
#                 action = "inserted"

#             SqlDb.commit()

#             if permit_id:
#                 SqlDb.execute_query(
#                     "UPDATE CommonHeaderTbl SET prmtStatus = 'AMD' WHERE PermitId = %s",
#                     [permit_id]
#                 )
#                 SqlDb.commit()

#         except Exception as e:
#             return Response({"error": f"Database error: {str(e)}"}, status=400)

#         return Response({"Result": f"Amend record {action} successfully", "MSGId": msg_id}, status=201)

class PostAmendTable(APIView):
    table = "CommonAmend"
    in_table = "InpaymentAmend"
    allowed_columns = {
        "Permitno", "AmendmentCount", "UpdateIndicator",
        "ReplacementPermitno", "DescriptionOfReason",
        "PermitExtension", "ExtendImportPeriod",
        "DeclarationIndigator", "AmendType",
        "TouchUser", "TouchTime", "MSGId",
    }

    # CommonAmend column -> InpaymentAmend column (only TouchTime differs)
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

            # ── 2. Mirror into InpaymentAmend (remapped TouchTime -> TouchTme) ──
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
            # NOTE: prmtStatus is NOT touched here anymore — it's already 'AMD'
            # from CopyAmend at the point the amend copy was created.

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response({"Result": f"Amend record {action} successfully", "MSGId": msg_id}, status=201)


# CANCEL PERMIT
# get Cancel Permit details by MSGId
class CancelPermit(APIView):
    table = "CommonCancel"
    def get(self, request):
        try:
            msg_id = request.GET.get("MSGId")
            if not msg_id:
                return Response({"error": "MSGId is required"}, status=status.HTTP_400_BAD_REQUEST)
            query = f"SELECT * FROM {self.table} WHERE MSGId = %s"
            rows = SqlDb.execute_query(query, [msg_id])
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

# # Post Cancel Permit
# class PostCancelPermit(APIView):
#     table = "CommonCancel"
#     allowed_columns = {
#             "Permitno", "UpdateIndicator", "ReplacementPermitno", "ReasonForCancel",
#             "DescriptionOfReason","DeclarationIndigator",
#             "TouchUser", "TouchTime", "MSGId", "CancelType"
#         }
#     def post(self, request):
#         item = request.data
#         if not item or not isinstance(item, dict):
#             return Response({"error": "No valid data provided"}, status=400)

#         item.pop("Id", None)
#         msg_id = item.get("MSGId", "")
#         permit_id = item.get("Permitno", "")

#         if not msg_id:
#             return Response({"error": "MSGId is required"}, status=400)

#         columns = [k for k in item.keys() if k in self.allowed_columns]
#         if not columns:
#             return Response({"error": "No valid columns provided"}, status=400)

#         try:
#             existing = SqlDb.execute_query(
#                 f"SELECT Id FROM {self.table} WHERE MSGId = %s", [msg_id]
#             )
#             if existing:
#                 update_cols = [c for c in columns if c != "MSGId"]
#                 set_clause = ", ".join([f"{c} = %s" for c in update_cols])
#                 values = [item[c] for c in update_cols] + [msg_id]
#                 SqlDb.execute_query(
#                     f"UPDATE {self.table} SET {set_clause} WHERE MSGId = %s", values
#                 )
#                 action = "updated"
#             else:
#                 col_str = ", ".join(columns)
#                 ph_str = ", ".join(["%s"] * len(columns))
#                 values = [item[c] for c in columns]
#                 SqlDb.execute_query(
#                     f"INSERT INTO {self.table} ({col_str}) VALUES ({ph_str})", values
#                 )
#                 action = "inserted"

#             SqlDb.commit()

#             if permit_id:
#                 SqlDb.execute_query(
#                     "UPDATE CommonHeaderTbl SET prmtStatus = 'CNL' WHERE PermitId = %s",
#                     [permit_id]
#                 )
#                 SqlDb.commit()

#         except Exception as e:
#             return Response({"error": f"Database error: {str(e)}"}, status=400)

#         return Response({"Result": f"Cancel record {action} successfully", "MSGId": msg_id}, status=201)

# Post Cancel Permit
class PostCancelPermit(APIView):
    table = "CommonCancel"
    in_table = "InpaymentCancel"
    allowed_columns = {
            "Permitno", "UpdateIndicator", "ReplacementPermitno", "ReasonForCancel",
            "DescriptionOfReason","DeclarationIndigator",
            "TouchUser", "TouchTime", "MSGId", "CancelType"

        }

    # CommonCancel column -> InpaymentCancel column
    # (InpaymentCancel has typo'd column names: ResonForCancel, TouchTme)
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

            # ── 2. Mirror into InpaymentCancel (remapped column names) ──
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
                SqlDb.execute_query(
                    "UPDATE InHeaderTbl SET prmtStatus = 'CNL' WHERE PermitId = %s",
                    [permit_id]
                )
                SqlDb.commit()

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response({"Result": f"Cancel record {action} successfully", "MSGId": msg_id}, status=201)

# Refund Permit
class RefundPermit(APIView):
    table = "CommonRefund"
    def get(self, request):
        try:
            msg_id = request.GET.get("MSGId")
            if not msg_id:
                return Response({"error": "MSGId is required"}, status=status.HTTP_400_BAD_REQUEST)
            query = f"SELECT * FROM {self.table} WHERE MSGId = %s"
            rows = SqlDb.execute_query(query, [msg_id])
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

# # Post Refund Permit
# class PostRefundPermit(APIView):
#     table = "CommonRefund"
#     allowed_columns = {
#         "Permitno","UpdateIndicator","ReplacementPermitno",
#         "TypeOfRefund","ReasonForRefund","DescriptionOfReason",  
#         "DeclarationIndigator","Additionalinfo","TotalGstAmt",
#         "TotalExciseAmt","TxtCusdutyAmt","TxtOtherAmt","RefundDatas","TouchUser","TouchTime","TouchTme",
#         "MSGId", 
#     }

#     def post(self, request):
#         item = request.data
#         if not item or not isinstance(item, dict):
#             return Response({"error": "No valid data provided"}, status=400)

#         item.pop("Id", None)
#         msg_id = item.get("MSGId", "")
#         permit_id = item.get("Permitno", "")

#         if not msg_id:
#             return Response({"error": "MSGId is required"}, status=400)

#         columns = [k for k in item.keys() if k in self.allowed_columns]
#         if not columns:
#             return Response({"error": "No valid columns provided"}, status=400)

#         try:
#             existing = SqlDb.execute_query(
#                 f"SELECT Id FROM {self.table} WHERE MSGId = %s", [msg_id]
#             )
#             if existing:
#                 update_cols = [c for c in columns if c != "MSGId"]
#                 set_clause = ", ".join([f"{c} = %s" for c in update_cols])
#                 values = [item[c] for c in update_cols] + [msg_id]
#                 SqlDb.execute_query(
#                     f"UPDATE {self.table} SET {set_clause} WHERE MSGId = %s", values
#                 )
#                 action = "updated"
#             else:
#                 col_str = ", ".join(columns)
#                 ph_str = ", ".join(["%s"] * len(columns))
#                 values = [item[c] for c in columns]
#                 SqlDb.execute_query(
#                     f"INSERT INTO {self.table} ({col_str}) VALUES ({ph_str})", values
#                 )
#                 action = "inserted"

#             SqlDb.commit()
#             saved = SqlDb.execute_query(
#             f"SELECT * FROM {self.table} WHERE MSGId = %s", [msg_id]
#             )
#             print("====== REFUND DATA ======")
#             print("TYPE:", type(saved))
#             print("RAW:", saved)
#             print("======== END DATA========")

#             if permit_id:
#                 SqlDb.execute_query(
#                     "UPDATE CommonHeaderTbl SET prmtStatus = 'RFD' WHERE PermitId = %s",
#                     [permit_id]
#                 )
#                 SqlDb.commit()


#         except Exception as e:
#             return Response({"error": f"Database error: {str(e)}"}, status=400)

#         return Response({"Result": f"Cancel record {action} successfully", "MSGId": msg_id}, status=201)

class PostRefundPermit(APIView):
    table = "CommonRefund"
    in_table = "InpaymentRefund"
    allowed_columns = {
        "Permitno","UpdateIndicator","ReplacementPermitno",
        "TypeOfRefund","ReasonForRefund","DescriptionOfReason",
        "DeclarationIndigator","Additionalinfo","TotalGstAmt",
        "TotalExciseAmt","TxtCusdutyAmt","TxtOtherAmt","RefundDatas","TouchUser","TouchTime","TouchTme",
        "MSGId",
    }

    IN_TABLE_COLUMN_MAP = {
        "Permitno": "Permitno",
        "UpdateIndicator": "UpdateIndicator",
        "ReplacementPermitno": "ReplacementPermitno",
        "TypeOfRefund": "TypeOfRefund",
        "ReasonForRefund": "ResonForREfund",
        "DescriptionOfReason": "DescriptionOfReason",
        "DeclarationIndigator": "DeclarationIndigator",
        "Additionalinfo": "Additionalinfo",
        "TouchUser": "TouchUser",
        "TouchTime": "TouchTme",
        "MSGId": "Msgid",
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
            # ── 1. CommonRefund upsert ──
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

            SqlDb.commit()

            # ── 2. Mirror into InpaymentRefund (mapped column names) ──
            in_item = {}
            for src_col, dst_col in self.IN_TABLE_COLUMN_MAP.items():
                if src_col in item:
                    in_item[dst_col] = item.get(src_col)

            if in_item:
                existing_in = SqlDb.execute_query(
                    f"SELECT Id FROM {self.in_table} WHERE Msgid = %s", [msg_id]
                )
                if existing_in:
                    in_update_cols = [c for c in in_item.keys() if c != "Msgid"]
                    if in_update_cols:
                        in_set_clause = ", ".join([f"{c} = %s" for c in in_update_cols])
                        in_values = [in_item[c] for c in in_update_cols] + [msg_id]
                        SqlDb.execute_query(
                            f"UPDATE {self.in_table} SET {in_set_clause} WHERE Msgid = %s",
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
                    "UPDATE CommonHeaderTbl SET prmtStatus = 'RFD', Status = 'NEW' WHERE PermitId = %s",
                    [permit_id]
                )
                SqlDb.execute_query(
                    "UPDATE InHeaderTbl SET prmtStatus = 'RFD', Status = 'NEW' WHERE PermitId = %s",
                    [permit_id]
                )
                SqlDb.commit()

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response({"Result": f"Refund record {action} successfully", "MSGId": msg_id}, status=201)
    
class RefundValSummary(APIView):
    table = "CommonRefundValSummary"
    def get(self, request):
        try:
            msg_id = request.GET.get("MSGId")
            if not msg_id:
                return Response({"error": "MSGId is required"}, status=status.HTTP_400_BAD_REQUEST)
            query = f"SELECT * FROM {self.table} WHERE MSGId = %s"
            rows = SqlDb.execute_query(query, [msg_id])
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

# class PostRefundValSummary(APIView):
#     table = "CommonRefundValSummary"
#     allowed_columns = {
#         "PermitId",
#         "totalgstAmt",
#         "totalexciseAmt",
#         "txtcusdutyAmt",
#         "txtotherAmt",
#         "MSGId"
#     }
#     def post(self, request):
#         try:
#             item = request.data
#             if not item or not isinstance(item, dict):
#                 return Response(
#                     {"error": "No valid data provided"},
#                     status=status.HTTP_400_BAD_REQUEST
#                 )
#             # item.pop("Id", None)
#             msg_id = item.get("MSGId")
#             permit_id = item.get("PermitId")
#             if not msg_id:
#                 return Response(
#                     {"error": "MSGId is required"},
#                     status=status.HTTP_400_BAD_REQUEST
#                 )
#             columns = [col for col in item.keys() if col in self.allowed_columns]
#             if not columns:
#                 return Response(
#                     {"error": "No valid columns provided"},
#                     status=status.HTTP_400_BAD_REQUEST
#                 )
#             check_query = f"SELECT MsgId FROM {self.table} WHERE MSGId = %s"
#             existing = SqlDb.execute_query(check_query, [msg_id])
#             if existing:
#                 update_columns = [col for col in columns if col != "MSGId"]

#                 if update_columns:
#                     set_clause = ", ".join([f"{col} = %s" for col in update_columns])
#                     values = [item[col] for col in update_columns]
#                     values.append(msg_id)

#                     update_query = f"""
#                         UPDATE {self.table}
#                         SET {set_clause}
#                         WHERE MSGId = %s
#                     """

#                     SqlDb.execute_query(update_query, values)

#                 action = "updated"

#             else:
#                 # INSERT
#                 col_str = ", ".join(columns)
#                 placeholders = ", ".join(["%s"] * len(columns))
#                 values = [item[col] for col in columns]

#                 insert_query = f"""
#                     INSERT INTO {self.table} ({col_str})
#                     VALUES ({placeholders})
#                 """

#                 SqlDb.execute_query(insert_query, values)

#                 action = "inserted"

#             # Update Header Status
#             if permit_id:
#                 SqlDb.execute_query(
#                     """
#                     UPDATE CommonHeaderTbl
#                     SET prmtStatus = 'RFD'
#                     WHERE PermitId = %s
#                     """,
#                     [permit_id]
#                 )

#             SqlDb.commit()

#             # Fetch Saved Data
#             saved = SqlDb.execute_query(
#                 f"SELECT * FROM {self.table} WHERE MSGId = %s",
#                 [msg_id]
#             )

#             return Response(
#                 {
#                     "message": f"Refund Summary {action} successfully",
#                     "data": saved
#                 },
#                 status=status.HTTP_200_OK
#             )

#         except Exception as e:
#             return Response(
#                 {"error": f"Database error: {str(e)}"},
#                 status=status.HTTP_400_BAD_REQUEST
#             )

class PostRefundValSummary(APIView):
    table = "CommonRefundValSummary"
    in_table = "refundvalsummary"
    allowed_columns = {
        "PermitId", "totalgstAmt", "totalexciseAmt",
        "txtcusdutyAmt", "txtotherAmt", "MSGId"
    }

    def post(self, request):
        try:
            item = request.data
            if not item or not isinstance(item, dict):
                return Response({"error": "No valid data provided"}, status=400)

            msg_id = item.get("MSGId")
            permit_id = item.get("PermitId")
            if not msg_id:
                return Response({"error": "MSGId is required"}, status=400)

            columns = [col for col in item.keys() if col in self.allowed_columns]
            if not columns:
                return Response({"error": "No valid columns provided"}, status=400)

            def upsert(table_name, key_col="MSGId"):
                existing = SqlDb.execute_query(
                    f"SELECT {key_col} FROM {table_name} WHERE {key_col} = %s", [msg_id]
                )
                if existing:
                    update_cols = [c for c in columns if c != "MSGId"]
                    if update_cols:
                        set_clause = ", ".join([f"{c} = %s" for c in update_cols])
                        values = [item[c] for c in update_cols] + [msg_id]
                        SqlDb.execute_query(
                            f"UPDATE {table_name} SET {set_clause} WHERE {key_col} = %s", values
                        )
                    return "updated"
                else:
                    col_str = ", ".join(columns)
                    placeholders = ", ".join(["%s"] * len(columns))
                    values = [item[c] for c in columns]
                    SqlDb.execute_query(
                        f"INSERT INTO {table_name} ({col_str}) VALUES ({placeholders})", values
                    )
                    return "inserted"

            # ── 1. CommonRefundValSummary ──
            action = upsert(self.table)

            # ── 2. Mirror into refundvalsummary (same column names) ──
            upsert(self.in_table)

            # ── 3. Update status on BOTH header tables ──
            if permit_id:
                SqlDb.execute_query(
                    "UPDATE CommonHeaderTbl SET prmtStatus = 'RFD', Status = 'NEW' WHERE PermitId = %s",
                    [permit_id]
                )
                SqlDb.execute_query(
                    "UPDATE InHeaderTbl SET prmtStatus = 'RFD', Status = 'NEW' WHERE PermitId = %s",
                    [permit_id]
                )

            SqlDb.commit()

            saved = SqlDb.execute_query(f"SELECT * FROM {self.table} WHERE MSGId = %s", [msg_id])
            return Response({"message": f"Refund Summary {action} successfully", "data": saved}, status=200)

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)


class GetReundItemSummByMsgId(APIView):
    table = "CommonReundItemSumm"
    def get(self, request):
        try:
            msg_id = request.GET.get("MSGId")
            if not msg_id:
                return Response({"error": "MSGId required"}, status=400)
            rows = SqlDb.execute_query(
                f"SELECT * FROM {self.table} WHERE MsgId = %s ORDER BY Sno",
                [msg_id]
            )
            return Response(rows, status=200)
        except Exception as e:
            return Response({"error": str(e)}, status=500)


# class PostReundItemSumm(APIView):
#     table = "CommonReundItemSumm"
#     allowed_columns = {
#         "PermitId", "ItemNo", "HsCode",
#         "TotalGstAmt", "TotalExciseAmt",
#         "TxtCusdutyAmt", "TxtOtherAmt",
#         "Sno", "MsgId"
#     }

#     def post(self, request):
#         item = request.data
#         if not item or not isinstance(item, dict):
#             return Response({"error": "No valid data provided"}, status=400)

#         item.pop("Id", None)
#         msg_id = item.get("MsgId", "")
#         sno = item.get("Sno", "")

#         if not msg_id:
#             return Response({"error": "MsgId is required"}, status=400)

#         columns = [k for k in item.keys() if k in self.allowed_columns]
#         if not columns:
#             return Response({"error": "No valid columns provided"}, status=400)

#         try:
#             existing = SqlDb.execute_query(
#                 f"SELECT Sno FROM {self.table} WHERE MsgId = %s AND Sno = %s",
#                 [msg_id, sno]
#             )
#             if existing:
#                 update_cols = [c for c in columns if c not in ["MsgId", "Sno"]]
#                 set_clause = ", ".join([f"{c} = %s" for c in update_cols])
#                 values = [item[c] for c in update_cols] + [msg_id, sno]
#                 SqlDb.execute_query(
#                     f"UPDATE {self.table} SET {set_clause} WHERE MsgId = %s AND Sno = %s",
#                     values
#                 )
#                 action = "updated"
#             else:
#                 col_str = ", ".join(columns)
#                 ph_str = ", ".join(["%s"] * len(columns))
#                 values = [item[c] for c in columns]
#                 SqlDb.execute_query(
#                     f"INSERT INTO {self.table} ({col_str}) VALUES ({ph_str})",
#                     values
#                 )
#                 action = "inserted"

#             SqlDb.commit()

#         except Exception as e:
#             return Response({"error": f"Database error: {str(e)}"}, status=400)

#         return Response(
#             {"Result": f"Refund item {action} successfully", "MsgId": msg_id},
#             status=201
#         )

class PostReundItemSumm(APIView):
    table = "CommonReundItemSumm"
    in_table = "reunditemsumm"
    allowed_columns = {
        "PermitId", "ItemNo", "HsCode",
        "TotalGstAmt", "TotalExciseAmt",
        "TxtCusdutyAmt", "TxtOtherAmt",
        "Sno", "MsgId"
    }

    # Maps CommonReundItemSumm column name -> reunditemsumm column name
    IN_TABLE_COLUMN_MAP = {
        "PermitId": "PermitId",
        "ItemNo": "ItemNo",
        "HsCode": "HsCode",
        "TotalGstAmt": "totalgstAmt",
        "TotalExciseAmt": "totalexciseAmt",
        "TxtCusdutyAmt": "txtcusdutyAmt",
        "TxtOtherAmt": "txtotherAmt",
        "Sno": "sno",
        "MsgId": "MSGId",
    }

    def post(self, request):
        item = request.data
        if not item or not isinstance(item, dict):
            return Response({"error": "No valid data provided"}, status=400)

        item.pop("Id", None)
        msg_id = item.get("MsgId", "")
        sno = item.get("Sno", "")
        permit_id = item.get("PermitId", "")

        if not msg_id:
            return Response({"error": "MsgId is required"}, status=400)

        columns = [k for k in item.keys() if k in self.allowed_columns]
        if not columns:
            return Response({"error": "No valid columns provided"}, status=400)

        try:
            # ── 1. CommonReundItemSumm upsert ──
            existing = SqlDb.execute_query(
                f"SELECT Sno FROM {self.table} WHERE MsgId = %s AND Sno = %s", [msg_id, sno]
            )
            if existing:
                update_cols = [c for c in columns if c not in ["MsgId", "Sno"]]
                set_clause = ", ".join([f"{c} = %s" for c in update_cols])
                values = [item[c] for c in update_cols] + [msg_id, sno]
                SqlDb.execute_query(
                    f"UPDATE {self.table} SET {set_clause} WHERE MsgId = %s AND Sno = %s", values
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

            # ── 2. Mirror into reunditemsumm (remapped, lowercase amount cols) ──
            in_item = {}
            for src_col, dst_col in self.IN_TABLE_COLUMN_MAP.items():
                if src_col in item:
                    in_item[dst_col] = item.get(src_col)

            if in_item:
                in_msg_id = in_item.get("MSGId", msg_id)
                in_sno = in_item.get("sno", sno)
                existing_in = SqlDb.execute_query(
                    f"SELECT sno FROM {self.in_table} WHERE MSGId = %s AND sno = %s",
                    [in_msg_id, in_sno]
                )
                if existing_in:
                    in_update_cols = [c for c in in_item.keys() if c not in ["MSGId", "sno"]]
                    if in_update_cols:
                        in_set_clause = ", ".join([f"{c} = %s" for c in in_update_cols])
                        in_values = [in_item[c] for c in in_update_cols] + [in_msg_id, in_sno]
                        SqlDb.execute_query(
                            f"UPDATE {self.in_table} SET {in_set_clause} WHERE MSGId = %s AND sno = %s",
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

            # ── 3. Update status on BOTH header tables ──
            if permit_id:
                SqlDb.execute_query(
                    "UPDATE CommonHeaderTbl SET prmtStatus = 'RFD', Status = 'NEW' WHERE PermitId = %s",
                    [permit_id]
                )
                SqlDb.execute_query(
                    "UPDATE InHeaderTbl SET prmtStatus = 'RFD', Status = 'NEW' WHERE PermitId = %s",
                    [permit_id]
                )

            SqlDb.commit()

        except Exception as e:
            return Response({"error": f"Database error: {str(e)}"}, status=400)

        return Response({"Result": f"Refund item {action} successfully", "MsgId": msg_id}, status=201)


class DownloadCcp(APIView):
    def get(self, request):
        data = request.GET.get("data", "")
        if not data:
            return HttpResponse("No permits provided", status=400)

        pdf_files = []
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        PDF_DIR = os.path.join(BASE_DIR, "DownloadCcp")
        os.makedirs(PDF_DIR, exist_ok=True)

        for PermitId in data.split(','):
            PermitId = PermitId.strip()
            if not PermitId:
                continue

            lftcol = 50
            rgtcol = 280
            countPage = 1

            # ── Fetch header data ──────────────────────────────────────────
            headers = SqlDb.execute_query(
                "SELECT * FROM CommonHeaderTbl WHERE PermitId = %s", [PermitId]
            )
            if not headers:
                continue
            Inheader = headers[0]

            importers = SqlDb.execute_query(
                "SELECT * FROM CommonImporter WHERE Code = %s",
                [Inheader['ImporterCompanyCode']]
            )
            Impo = importers[0] if importers else {}

            loadingPorts = SqlDb.execute_query(
                "SELECT * FROM LoadingPort WHERE portcode = %s",
                [Inheader['LoadingPortCode']]
            )
            loadingPort = loadingPorts[0] if loadingPorts else {}

            try:
                inwardCarriers = SqlDb.execute_query(
                    "SELECT * FROM CommonInwardCarrierAgent WHERE Code = %s",
                    [Inheader['InwardCarrierAgentCode']]
                )
                InwardCarrier = inwardCarriers[0] if inwardCarriers else {}
                InwardCarrierName = (InwardCarrier.get('Name', '') or '') + (InwardCarrier.get('Name1', '') or '')
            except:
                InwardCarrierName = ""

            Cpcs = SqlDb.execute_query(
                "SELECT DISTINCT CpcType FROM CommonCPCDtl WHERE PermitId = %s",
                [PermitId]
            )

            declarants = SqlDb.execute_query(
                "SELECT * FROM DeclarantCompany WHERE TradeNetMailboxID = %s",
                [Inheader['TradeNetMailboxID']]
            )
            Declarant = declarants[0] if declarants else {}

            inpstrtdate = ""
            inpEndate = ""
            try:
                pmts = SqlDb.execute_query(
                    "SELECT * FROM CommonPMT WHERE PermitNumber = %s",
                    [Inheader['PermitNumber']]
                )
                if pmts:
                    pmt0 = pmts[0]
                    if pmt0.get('StartDate'):
                        inpstrtdate = datetime.strptime(str(pmt0['StartDate']), '%Y-%m-%d').strftime('%d/%m/%Y')
                    if pmt0.get('EndDate'):
                        inpEndate = datetime.strptime(str(pmt0['EndDate']), '%Y-%m-%d').strftime('%d/%m/%Y')
            except:
                pass

            importerName = (Impo.get('Name', '') or '') + (Impo.get('Name1', '') or '')

            buffer = io.BytesIO()
            p = canvas.Canvas(buffer, pagesize=(595, 841))

            PermitNumber = (Inheader.get('PermitNumber') or '').strip()
            if PermitNumber == "":
                draft = "DRAFT"
                inpstrtdate = ""
                inpEndate = ""
            else:
                draft = PermitNumber
                barcode = code39.Standard39(
                    draft, barHeight=36.0, barWidth=1.1, baseline=9.0,
                    size=12.0, N=3.0, X=1.0, StartsStopText=False, Extended=False
                )
                barcode.drawOn(p, 330, 760)

            # ── Page 1 ─────────────────────────────────────────────────────
            p.setFont('Courier-Bold', 10)
            p.drawString(480, 750, draft)
            p.setFont('Courier', 10)
            p.drawString(400, 750, "PERMIT NO : ")
            p.drawString(rgtcol, 700, "CARGO CLEARANCE PERMIT")
            p.drawString(460, 700, f"PG : {countPage} OF")
            countPage += 1

            p.drawString(lftcol, 670, "MESSAGE TYPE      : IN-PAYMENT PERMIT")
            decltype = str(Inheader.get('DeclarationType', ''))[6:].upper() if Inheader.get('DeclarationType') else ''
            p.drawString(lftcol, 660, "DECLARATION TYPE  : " + decltype)

            p.drawString(lftcol, 630, "IMPORTER:")
            if len(importerName) >= 35:
                p.drawString(lftcol, 620, importerName[:35].upper())
                p.drawString(lftcol, 610, importerName[35:].upper())
            else:
                p.drawString(lftcol, 620, importerName.upper())

            p.drawString(lftcol, 600, (Impo.get('CRUEI', '') or '').upper())
            p.drawString(lftcol, 590, "EXPORTER:")
            p.drawString(lftcol, 580, " ")
            p.drawString(lftcol, 570, " ")
            p.drawString(lftcol, 560, " ")
            p.drawString(lftcol, 550, "HANDLING AGENT: ")
            p.drawString(lftcol, 540, " ")
            p.drawString(lftcol, 530, " ")
            p.drawString(lftcol, 520, " ")
            p.drawString(lftcol, 510, " ")
            p.drawString(lftcol, 500, "PORT OF LOADING/NEXT PORT OF CALL:")
            p.drawString(lftcol, 490, (loadingPort.get('portname', '') or '').upper())
            p.drawString(lftcol, 480, "PORT OF DISCHARGE/FINAL PORT OF CALL ")
            p.drawString(lftcol, 470, " ")
            p.drawString(lftcol, 460, "COUNTRY OF FINAL DESTINATION:")
            p.drawString(lftcol, 450, " ")
            p.drawString(lftcol, 440, "INWARD CARRIER AGENT: ")

            if len(InwardCarrierName) >= 35:
                p.drawString(lftcol, 430, InwardCarrierName[:35].upper())
                p.drawString(lftcol, 420, InwardCarrierName[35:70].upper())
                p.drawString(lftcol, 410, InwardCarrierName[70:].upper())
            else:
                p.drawString(lftcol, 430, InwardCarrierName.upper())

            p.drawString(lftcol, 400, "OUTWARD CARRIER AGENT: ")
            p.drawString(lftcol, 390, " ")
            p.drawString(lftcol, 380, " ")
            p.drawString(lftcol, 370, " ")
            p.drawString(lftcol, 360, "PLACE OF RELEASE: ")

            rely = 350
            relVal = (Inheader.get('ReleaseLocName', '') or '').upper()
            chr = 0
            if len(relVal) >= 32:
                for i in range(len(relVal)):
                    if (i + 1) % 32 == 0:
                        p.drawString(lftcol, rely, relVal[chr:i])
                        chr = i + 1
                        rely -= 10
                p.drawString(lftcol, rely, relVal[chr:])
                rely -= 100
            else:
                p.drawString(lftcol, rely, relVal)
                rely -= 10

            p.drawString(lftcol, rely, Inheader.get('ReleaseLocation', '') or '')

            # Licence lines
            p.drawString(lftcol, 250, "LICENCE NO:")
            licence = ((Inheader.get('License', '') or '').upper()).split(',')
            for li_idx, li_y in enumerate([240, 230, 220, 210, 200]):
                p.drawString(lftcol, li_y, licence[li_idx] if li_idx < len(licence) else '')

            p.drawString(lftcol, 50, "-" * 80)
            MSGId = Inheader.get('MSGId', '') or ''
            p.drawString(lftcol, 40, f"UNIQUE REF : {(Declarant.get('CRUEI','') or '').upper()} {MSGId[:8].upper()} {MSGId[8:].upper()}")

            # ── Right column page 1 ────────────────────────────────────────
            rgy = 630
            p.drawString(rgtcol, rgy, f"VALIDITY PERIOD      : {inpstrtdate} - ")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"                       {inpEndate}")
            rgy -= 20

            totalGross = "{:.3f}".format(float(Inheader.get('TotalGrossWeight') or 0))
            p.drawString(rgtcol, rgy, f"TOTAL GROSS WT/UNIT  : {totalGross:>18}/{Inheader.get('TotalGrossWeightUOM','')}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOTAL OUTER PACK/UNIT: {str(Inheader.get('TotalOuterPack','') or ''):>18}/{Inheader.get('TotalOuterPackUOM','')}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOT EXCISE DUT PAYABLE  : S${str(Inheader.get('TotalExDutyAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOT CUSTOMS DUT PAYABLE : S${str(Inheader.get('TotalCusDutyAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOT OTHER TAX PAYABLE   : S${str(Inheader.get('TotalODutyAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOTAL GST AMT           : S${str(Inheader.get('TotalGSTTaxAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOTAL AMOUNT PAYABLE    : S${str(Inheader.get('TotalAmtPay','') or ''):>17}")
            rgy -= 10
            cargoPack = str(Inheader.get('CargoPackType', '') or '')[3:].upper()
            p.drawString(rgtcol, rgy, f"CARGO PACKING TYPE: {cargoPack} ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "IN TRANSPORT IDENTIFIER: ")
            rgy -= 10

            inwardMode = str(Inheader.get('InwardTransportMode', '') or '')[4:].upper()
            if inwardMode in ("SEA", "AIR"):
                transId = Inheader.get('VesselName', '') or ''
            else:
                transId = Inheader.get('TransportId', '') or ''

            if inwardMode == "SEA":
                CoveyanceNo = Inheader.get('VoyageNumber', '') or ''
            elif inwardMode == "AIR":
                CoveyanceNo = Inheader.get('FlightNO', '') or ''
            else:
                CoveyanceNo = Inheader.get('ConveyanceRefNo', '') or ''

            p.drawString(rgtcol, rgy, transId.upper())
            rgy -= 10
            p.drawString(rgtcol, rgy, f"CONVEYANCE REFERENCE NO: {CoveyanceNo.upper()} ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "OBL/MAWB NO: ")
            rgy -= 10
            p.drawString(rgtcol, rgy, Inheader.get('OceanBillofLadingNo', '') or '')
            rgy -= 10
            arrDate = Inheader.get('ArrivalDate')
            arrDateStr = arrDate.strftime('%d/%m/%Y') if arrDate else ""
            p.drawString(rgtcol, rgy, f"ARRIVAL DATE         : {arrDateStr}")
            rgy -= 10
            p.drawString(rgtcol, rgy, "OU TRANSPORT IDENTIFIER: ")
            rgy -= 10
            p.drawString(rgtcol, rgy, " ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "CONVEYANCE REFERENCE NO:  ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "OBL/MAWB/UCR NO: ")
            rgy -= 10
            p.drawString(rgtcol, rgy, " ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "DEPARTURE DATE       : ")
            rgy -= 20
            p.drawString(rgtcol, rgy, "CERTIFICATE NO:  ")
            rgy -= 30
            p.drawString(rgtcol, rgy, "PLACE OF RECEIPT:")
            rgy -= 10

            rely = rgy
            recVal = Inheader.get('RecepitLocName', '') or ''
            chr1 = 0
            if len(recVal) >= 32:
                for i in range(len(recVal)):
                    if (i + 1) % 32 == 0:
                        p.drawString(rgtcol, rely, recVal[chr1:i])
                        chr1 = i + 1
                        rely -= 10
                p.drawString(rgtcol, rely, recVal[chr1:])
                rely -= 10
            else:
                p.drawString(rgtcol, rely, recVal)
                rely -= 10

            p.drawString(rgtcol, rely, Inheader.get('RecepitLocation', '') or '')
            p.drawString(rgtcol, 250, "CUSTOMS PROCEDURE CODE (CPC) : ")
            rely = 240
            for cpc_row in Cpcs:
                p.drawString(rgtcol, rely, cpc_row.get('CpcType', '') or '')
                rely -= 10
            if str(Inheader.get('Cnb', '') or '').lower() == 'true':
                p.drawString(rgtcol, rely, "CNB")
                rely -= 10

            p.showPage()
            # ── Page 1 complete ────────────────────────────────────────────

            # ── Items pages ────────────────────────────────────────────────
            snox = 50
            hscodex = 100
            currentx = 180
            prviousx = 320
            makingx = 50
            cityx = 120
            brandx = 220
            itemy = 820

            ItemDtls = SqlDb.execute_query(
                "SELECT * FROM CommonItemDtl WHERE PermitId = %s ORDER BY ItemNo",
                [PermitId]
            )
            InvoiceDtls = SqlDb.execute_query(
                "SELECT * FROM CommonInvoiceDtl WHERE PermitId = %s",
                [PermitId]
            )
            ContainerDtls = SqlDb.execute_query(
                "SELECT * FROM CommonContainerDtl WHERE PermitId = %s",
                [PermitId]
            )

            def itemyF(itemy):
                nonlocal countPage, p
                if itemy <= 70:
                    itemy = 820
                    p.showPage()
                    p.setFont('Courier', 10)
                else:
                    itemy -= 10
                if 700 <= itemy <= 820:
                    p.setFont('Courier', 10)
                    p.drawString(rgtcol, itemy, "CARGO CLEARANCE PERMIT ")
                    p.drawString(460, itemy, f"PG : {countPage} OF ")
                    countPage += 1
                    itemy -= 10
                    p.drawString(lftcol, itemy, "PERMIT NO : " + draft)
                    p.drawString(rgtcol, itemy, "======================")
                    itemy -= 10
                    p.drawString(rgtcol, itemy, "(CONTINUATION PAGE)")
                    itemy -= 20
                    p.drawString(lftcol, itemy, "CONSIGNMENT DETAILS")
                    itemy -= 10
                    p.drawString(lftcol, itemy, "-" * 80)
                    itemy -= 10
                    p.drawString(snox, itemy, "S/NO ")
                    p.drawString(hscodex, itemy, "HS CODE")
                    p.drawString(currentx, itemy, "CURRENT LOT NO")
                    p.drawString(prviousx, itemy, "PREVIOUS LOT NO")
                    itemy -= 10
                    p.drawString(makingx, itemy, "MARKING")
                    p.drawString(cityx, itemy, "CTY OF ORIGIN")
                    p.drawString(brandx, itemy, "BRAND NAME")
                    p.drawString(prviousx, itemy, "MODEL")
                    itemy -= 10
                    for itd in ItemDtls:
                        if itd.get('InHAWBOBL', ''):
                            p.drawString(lftcol, itemy, "IN HAWB/HUCR/HBL")
                            p.drawString(prviousx, itemy, "OUT HAWB/HUCR/HBL")
                            itemy -= 10
                            break
                    p.drawString(lftcol, itemy, "PACKING/GOODS DESCRIPTION")
                    p.drawString(prviousx, itemy, "HS QUANTITY & UNIT ")
                    itemy -= 10
                    p.drawString(prviousx, itemy, "CIF/FOB VALUE (S$) ")
                    itemy -= 10
                    for itd in ItemDtls:
                        if float(itd.get('LSPValue') or 0) != 0.0:
                            p.drawString(prviousx, itemy, "LSP VALUE (S$) ")
                            itemy -= 10
                            break
                    p.drawString(prviousx, itemy, 'GST AMOUNT (S$)')
                    itemy -= 10
                    for itd in ItemDtls:
                        if float(itd.get('TotalDutiableQty') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'DUT QTY/WT/VOL & UNIT')
                            itemy -= 10
                            break
                    for itd in ItemDtls:
                        if float(itd.get('ExciseDutyAmount') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'EXCISE DUTY PAYABLE (S$)')
                            itemy -= 10
                            break
                    for itd in ItemDtls:
                        if float(itd.get('CustomsDutyAmount') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'CUSTOMS DUTY PAYABLE(S$)')
                            itemy -= 10
                            break
                    for itd in ItemDtls:
                        if float(itd.get('OtherTaxAmount') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'OTHER TAX PAYABLE(S$) ')
                            itemy -= 10
                            break
                    p.drawString(snox, itemy, "MANUFACTURER'S NAME ")
                    itemy -= 10
                    p.drawString(snox, itemy, '-' * 79)
                    itemy -= 10
                    p.drawString(lftcol, 50, "-" * 80)
                    p.drawString(lftcol, 40, f"UNIQUE REF : {(Declarant.get('CRUEI','') or '').upper()} {MSGId[:8].upper()} {MSGId[8:].upper()}")
                return itemy

            for item in ItemDtls:
                itemy = itemyF(itemy)
                p.drawString(snox, itemy, f"   {int(item.get('ItemNo',0)):02d}")
                p.drawString(hscodex, itemy, str(item.get('HSCode', '') or ''))
                p.drawString(currentx, itemy, (item.get('CurrentLot', '') or '').upper())
                p.drawString(prviousx, itemy, (item.get('PreviousLot', '') or '').upper())
                itemy = itemyF(itemy)
                making = (item.get('Making', '') or '').upper()
                if '--SELECT--' != making:
                    p.drawString(makingx, itemy, making[:2])
                p.drawString(cityx - 40, itemy, (item.get('Contry', '') or '').upper())
                p.drawString(brandx - 105, itemy, (item.get('Brand', '') or '').upper())
                p.drawString(prviousx, itemy, (item.get('Model', '') or '').upper())
                itemy = itemyF(itemy)
                if item.get('InHAWBOBL', ''):
                    p.drawString(lftcol, itemy, (item.get('InHAWBOBL', '') or '').upper())
                    itemy = itemyF(itemy)

                ItemDescr = str(item.get('Description', '') or '').upper()

                def write_val_right(val, y_pos, is_qty=False, uom=''):
                    try:
                        num = float(val or 0)
                        dec = str(val).split('.')
                        dec_part = dec[1] if len(dec) > 1 else '0'
                        if is_qty:
                            p.drawString(prviousx + 100, y_pos, f'{int(num):8d}.{dec_part} {uom}')
                        else:
                            p.drawString(prviousx + 100, y_pos, f'{int(num):10d}.{dec_part}')
                    except:
                        pass

                if '\n' in ItemDescr:
                    itemDesc = ItemDescr.split('\n')
                    def get_desc(i): return itemDesc[i].replace('\r', '') if i < len(itemDesc) else ''

                    p.drawString(lftcol, itemy, get_desc(0))
                    write_val_right(item.get('HSQty'), itemy, True, item.get('HSUOM',''))
                    itemy = itemyF(itemy)

                    p.drawString(lftcol, itemy, get_desc(1))
                    write_val_right(item.get('CIFFOB'), itemy)
                    itemy = itemyF(itemy)

                    if float(item.get('LSPValue') or 0) != 0.0:
                        p.drawString(lftcol, itemy, get_desc(2))
                        write_val_right(item.get('LSPValue'), itemy)
                        itemy = itemyF(itemy)
                    else:
                        p.drawString(lftcol, itemy, get_desc(2))

                    write_val_right(item.get('GSTAmount'), itemy)
                    itemy = itemyF(itemy)

                    if float(item.get('TotalDutiableQty') or 0) != 0.0:
                        p.drawString(lftcol, itemy, get_desc(3))
                        write_val_right(item.get('TotalDutiableQty'), itemy, True, item.get('TotalDutiableUOM',''))
                        itemy = itemyF(itemy)
                    elif len(itemDesc) >= 4:
                        p.drawString(lftcol, itemy, get_desc(3))
                        itemy = itemyF(itemy)

                    for duty_field, desc_idx in [
                        ('ExciseDutyAmount', 4), ('CustomsDutyAmount', 5), ('OtherTaxAmount', 6)
                    ]:
                        if float(item.get(duty_field) or 0) != 0.0:
                            p.drawString(lftcol, itemy, get_desc(desc_idx))
                            write_val_right(item.get(duty_field), itemy)
                            itemy = itemyF(itemy)
                        elif len(itemDesc) >= desc_idx + 1:
                            p.drawString(lftcol, itemy, get_desc(desc_idx))
                            itemy = itemyF(itemy)
                else:
                    # Flat description mode
                    slices = [(0,50),(50,100),(100,150),(150,200),(200,250),(250,300),(300,350)]
                    fields = [
                        ('HSQty', True, item.get('HSUOM','')),
                        ('CIFFOB', False, ''),
                        ('LSPValue', False, ''),
                        ('GSTAmount', False, ''),
                        ('TotalDutiableQty', True, item.get('TotalDutiableUOM','')),
                        ('ExciseDutyAmount', False, ''),
                        ('CustomsDutyAmount', False, ''),
                        ('OtherTaxAmount', False, ''),
                    ]
                    duty_names = ['ExciseDutyAmount','CustomsDutyAmount','OtherTaxAmount']
                    for idx, (s, e) in enumerate(slices):
                        if len(ItemDescr) >= s or idx == 0:
                            p.drawString(lftcol, itemy, ItemDescr[s:e])
                            if idx < len(fields):
                                fname, is_qty, uom = fields[idx]
                                val = item.get(fname)
                                if fname == 'LSPValue' and float(val or 0) == 0.0:
                                    pass
                                elif val is not None:
                                    write_val_right(val, itemy, is_qty, uom)
                            itemy = itemyF(itemy)

                for ext_slice in [(350,400),(400,450),(450,500),(500,None)]:
                    s, e = ext_slice
                    chunk = ItemDescr[s:e] if e else ItemDescr[s:]
                    if len(ItemDescr) >= s and chunk:
                        p.drawString(lftcol, itemy, chunk)
                        itemy = itemyF(itemy)

                # Manufacturer name
                for inv in InvoiceDtls:
                    if inv.get('InvoiceNo') == item.get('InvoiceNo'):
                        manfs = SqlDb.execute_query(
                            "SELECT * FROM CommonSupplierManufacturerPart WHERE Code = %s",
                            [inv.get('SupplierCode')]
                        )
                        if manfs:
                            manf = manfs[0]
                            manfName = manf.get('Name', '') or ''
                            p.drawString(lftcol, itemy, manfName[:50])
                            itemy = itemyF(itemy)
                            if len(manfName) >= 50:
                                p.drawString(lftcol, itemy, manfName[50:])

                # CASC details
                Cascdtls = SqlDb.execute_query(
                    "SELECT * FROM CommonCASCDtl WHERE PermitId = %s AND ItemNo = %s",
                    [PermitId, item.get('ItemNo')]
                )
                if Cascdtls:
                    for casc_sno, casc_id_key in [('01','Casc1'),('02','Casc2'),('03','Casc3'),('04','Casc4'),('05','Casc5')]:
                        matched = [c for c in Cascdtls if c.get('CASCId') == casc_id_key]
                        if matched:
                            casc = matched[0]
                            p.drawString(lftcol, itemy, "-" * 80)
                            itemy = itemyF(itemy)
                            p.drawString(lftcol, itemy, 'S/NO')
                            p.drawString(lftcol + 70, itemy, 'CA/SC PRODUCT CODE ')
                            p.drawString(lftcol + 280, itemy, 'CA/SC PRODUCT QTY & UNIT')
                            itemy = itemyF(itemy)
                            p.drawString(lftcol, itemy, f'   {casc_sno}')
                            p.drawString(lftcol + 70, itemy, (casc.get('ProductCode', '') or '').upper())
                            qty = casc.get('Quantity')
                            if qty and str(qty) != "0.0000":
                                try:
                                    dec = str(qty).split('.')
                                    p.drawString(lftcol + 280, itemy, f"{int(float(qty)):10d}.{dec[1] if len(dec)>1 else '0'} {casc.get('ProductUOM','')}")
                                except:
                                    pass
                            itemy = itemyF(itemy)
                            p.drawString(lftcol, itemy, "-" * 80)
                            break

                # Engine Capacity
                engCap = str(item.get('EngineCapcity', '') or '')
                if engCap not in ("0.00", ""):
                    p.drawString(snox, itemy, '-' * 79)
                    itemy = itemyF(itemy)
                    p.drawString(lftcol, itemy, 'S/NO')
                    p.drawString(lftcol + 70, itemy, 'ENGINE NO/CHASSIS NO ')
                    itemy = itemyF(itemy)
                    p.drawString(lftcol, itemy, f"   {int(item.get('ItemNo',0)):02d}")
                    p.drawString(lftcol + 70, itemy, engCap)
                    itemy = itemyF(itemy)
                    p.drawString(snox, itemy, '-' * 79)

                p.drawString(snox, itemy, '-' * 79)
                itemy = itemyF(itemy) - 20
                p.drawString(snox, itemy, '-' * 79)

            # ── itemyF1 for trailing sections ──────────────────────────────
            def itemyF1(itemy):
                nonlocal countPage, p
                if itemy <= 60:
                    itemy = 820
                    p.showPage()
                    p.setFont('Courier', 10)
                else:
                    itemy -= 10
                if 780 <= itemy <= 820:
                    p.setFont('Courier', 10)
                    p.drawString(rgtcol, itemy, "CARGO CLEARANCE PERMIT ")
                    p.drawString(460, itemy, f"PG : {countPage} OF")
                    countPage += 1
                    itemy -= 10
                    p.drawString(lftcol, itemy, "PERMIT NO : " + draft)
                    p.drawString(rgtcol, itemy, "======================")
                    itemy -= 10
                    p.drawString(rgtcol, itemy, "(CONTINUATION PAGE)")
                    itemy -= 20
                    p.drawString(lftcol, itemy, "CONSIGNMENT DETAILS")
                    itemy -= 10
                    p.drawString(lftcol, itemy, "-" * 80)
                    itemy -= 10
                    p.drawString(lftcol, 50, "-" * 80)
                    p.drawString(lftcol, 40, f"UNIQUE REF : {(Declarant.get('CRUEI','') or '').upper()} {MSGId[:8].upper()} {MSGId[8:].upper()}")
                return itemy

            # Trader remarks
            tradeRemarks = str(Inheader.get('TradeRemarks', '') or '')
            if tradeRemarks:
                itemy = itemyF1(itemy)
                p.drawString(lftcol, itemy, "TRADER'S REMARKS")
                itemy = itemyF1(itemy)
                TradeRe = tradeRemarks.upper()
                if '\n' in TradeRe:
                    for trd in TradeRe.split('\n'):
                        p.drawString(lftcol, itemy, trd)
                        itemy = itemyF1(itemy)
                else:
                    for s in range(0, max(len(TradeRe), 1), 80):
                        chunk = TradeRe[s:s+80]
                        if chunk:
                            p.drawString(lftcol, itemy, chunk)
                            itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, '-' * 79)

            # Containers
            if ContainerDtls:
                itemy = itemyF1(itemy)
                p.drawString(lftcol, itemy, "CONTAINER IDENTIFIERS")
                itemy = itemyF1(itemy)
                for cont in ContainerDtls:
                    p.drawString(lftcol, itemy, f"   {int(cont.get('RowNo',0)):02d})")
                    p.drawString(lftcol + 50, itemy, str(cont.get('ContainerNo', '') or '').upper())
                    size = str(cont.get('Size', '') or '')
                    p.drawString(lftcol + 130, itemy, f"{size[:3]}  {size[3:5]}")
                    p.drawString(lftcol + 200, itemy, str(cont.get('Weight', '') or '')[:3])
                    p.drawString(lftcol + 240, itemy, str(cont.get('SealNo', '') or ''))
                    itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, '-' * 79)
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, "NO UNAUTHORISED ADDITION/AMENDMENT TO THIS PERMIT MAY BE MADE AFTER APPROVAL")
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, '-' * 79)
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, "NAME OF COMPANY:")
            decName = Declarant.get('name', '') or ''
            p.drawString(lftcol + 110, itemy, decName[:67])
            itemy = itemyF1(itemy)
            p.drawString(lftcol + 110, itemy, decName[67:])
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, "DECLARANT NAME :")
            decPersonName = Declarant.get('DeclarantName', '') or ''
            p.drawString(lftcol + 110, itemy, decPersonName[:67])
            itemy = itemyF1(itemy)
            p.drawString(lftcol + 110, itemy, decPersonName[67:])
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, "DECLARANT CODE :")
            decCode = Declarant.get('DeclarantCode', '') or ''
            p.drawString(lftcol + 110, itemy, "XXXX" + decCode[-5:])
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, "TEL NO         : ")
            p.drawString(lftcol + 110, itemy, Declarant.get('DeclarantTel', '') or '')
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, '-' * 79)
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, "CONTROLLING AGENCY/CUSTOMS CONDITIONS ")
            itemy = itemyF1(itemy)

            for pmt in pmts:
                p.setFont('Courier-Bold', 10)
                p.drawString(lftcol, itemy, pmt.get('ConditionCode', '') or '')
                p.drawString(lftcol + 30, itemy, "-")
                p.setFont('Courier', 10)
                condDesc = pmt.get('ConditionDesc', '') or pmt.get('ConditionDescription', '') or ''
                p.drawString(lftcol + 40, itemy, condDesc[:73])
                itemy = itemyF1(itemy)
                for s in range(73, len(condDesc), 80):
                    p.drawString(lftcol, itemy, condDesc[s:s+80])
                    itemy = itemyF1(itemy)

            # ── Save PDF ───────────────────────────────────────────────────
            safe_draft = draft if draft != "DRAFT" else f"Draft_{MSGId.strip()}"
            filename = os.path.join(PDF_DIR, f"1{safe_draft}.pdf")
            final_filename = os.path.join(PDF_DIR, f"{safe_draft}.pdf")

            p.setTitle(safe_draft)
            p.save()

            with open(filename, 'wb') as f:
                f.write(buffer.getbuffer())
            buffer.close()

            # ── Merge total-page-count overlay ─────────────────────────────
            with open(filename, "rb") as pdf_file:
                existing_pdf = PdfReader(pdf_file)
                total_pages = len(existing_pdf.pages)
                output = PdfWriter()

                packet2 = io.BytesIO()
                can2 = canvas.Canvas(packet2, pagesize=(595, 841))
                can2.setFont('Courier', 10)
                can2.drawString(520, 700, str(total_pages))
                can2.showPage()
                can2.setFont('Courier', 10)
                can2.drawString(520, 810, str(total_pages))
                can2.showPage()
                can2.setFont('Courier', 10)
                can2.drawString(520, 820, str(total_pages))
                can2.save()
                packet2.seek(0)
                new_pdf = PdfReader(packet2)

                for i in range(total_pages):
                    page = existing_pdf.pages[i]
                    if i == 0:
                        page.merge_page(new_pdf.pages[0])
                    elif i == 1:
                        page.merge_page(new_pdf.pages[1])
                    else:
                        page.merge_page(new_pdf.pages[2])
                    output.add_page(page)

                with open(final_filename, "wb") as out_stream:
                    output.write(out_stream)

            packet2.close()
            pdf_files.append(final_filename)
            if os.path.exists(filename):
                os.remove(filename)

        # ── Return ZIP ─────────────────────────────────────────────────────
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, mode='w') as zf:
            for fpath in pdf_files:
                if os.path.isfile(fpath):
                    zf.write(fpath, arcname=os.path.join("Files", os.path.basename(fpath)))

        zip_buffer.seek(0)
        response = HttpResponse(zip_buffer, content_type='application/zip')
        response['Content-Disposition'] = (
            f'attachment; filename="Zip_{datetime.now().strftime("%Y-%B-%d-%H%M")}.zip"'
        )
        return response

# PRINT CCP
class PrintCcp(APIView):
    def get(self, request, permit_id):
        try:
            lftcol = 50
            rgtcol = 280
            countPage = 1

            # ── Fetch header ──────────────────────────────────────────────
            headers = SqlDb.execute_query(
                "SELECT * FROM CommonHeaderTbl WHERE PermitId = %s", [permit_id]
            )
            if not headers:
                return HttpResponse("Permit not found", status=404)
            Inheader = headers[0]

            # ── Fetch importer ────────────────────────────────────────────
            importers = SqlDb.execute_query(
                "SELECT * FROM CommonImporter WHERE Code = %s",
                [Inheader.get('ImporterCompanyCode', '')]
            )
            Impo = importers[0] if importers else {}

            # ── Fetch loading port ────────────────────────────────────────
            loadingPorts = SqlDb.execute_query(
                "SELECT * FROM LoadingPort WHERE portcode = %s",
                [Inheader.get('LoadingPortCode', '')]
            )
            loadingPort = loadingPorts[0] if loadingPorts else {}

            # ── Fetch inward carrier agent ────────────────────────────────
            try:
                inwardCarriers = SqlDb.execute_query(
                    "SELECT * FROM CommonInwardCarrierAgent WHERE Code = %s",
                    [Inheader.get('InwardCarrierAgentCode', '')]
                )
                InwardCarrier = inwardCarriers[0] if inwardCarriers else {}
                InwardCarrierName = (InwardCarrier.get('Name', '') or '') + (InwardCarrier.get('Name1', '') or '')
            except Exception:
                InwardCarrierName = ""

            # ── Fetch distinct CPC types ──────────────────────────────────
            Cpcs = SqlDb.execute_query(
                "SELECT DISTINCT CPCType FROM CommonCPCDtl WHERE PermitId = %s",
                [permit_id]
            )

            # ── Fetch declarant ───────────────────────────────────────────
            declarants = SqlDb.execute_query(
                "SELECT * FROM DeclarantCompany WHERE TradeNetMailboxID = %s",
                [Inheader.get('TradeNetMailboxID', '')]
            )
            Declarant = declarants[0] if declarants else {}

            # ── Fetch PMT validity dates ──────────────────────────────────
            inpstrtdate = ""
            inpEndate = ""
            pmts = []
            try:
                pmts = SqlDb.execute_query(
                    "SELECT * FROM CommonPMT WHERE PermitNumber = %s ORDER BY Sno",
                    [Inheader.get('PermitNumber', '')]
                )
                if pmts:
                    pmt0 = pmts[0]
                    if pmt0.get('StartDate'):
                        inpstrtdate = datetime.strptime(str(pmt0['StartDate']), '%Y-%m-%d').strftime('%d/%m/%Y')
                    if pmt0.get('EndDate'):
                        inpEndate = datetime.strptime(str(pmt0['EndDate']), '%Y-%m-%d').strftime('%d/%m/%Y')
            except Exception:
                pass

            importerName = (Impo.get('Name', '') or '') + (Impo.get('Name1', '') or '')
            MSGId = Inheader.get('MSGId', '') or ''

            # ── Fetch child records ───────────────────────────────────────
            ItemDtls = SqlDb.execute_query(
                "SELECT * FROM CommonItemDtl WHERE PermitId = %s ORDER BY ItemNo",
                [permit_id]
            )
            InvoiceDtls = SqlDb.execute_query(
                "SELECT * FROM CommonInvoiceDtl WHERE PermitId = %s",
                [permit_id]
            )
            ContainerDtls = SqlDb.execute_query(
                "SELECT * FROM CommonContainerDtl WHERE PermitId = %s",
                [permit_id]
            )

            # ── Build PDF ─────────────────────────────────────────────────
            buffer = io.BytesIO()
            p = canvas.Canvas(buffer, pagesize=(595, 841))

            PermitNumber = (Inheader.get('PermitNumber') or '').strip()
            if PermitNumber == "":
                draft = "DRAFT"
                inpstrtdate = ""
                inpEndate = ""
            else:
                draft = PermitNumber
                barcode = code39.Standard39(
                    draft, barHeight=36.0, barWidth=1.1, baseline=9.0,
                    size=12.0, N=3.0, X=1.0, StartsStopText=False, Extended=False
                )
                barcode.drawOn(p, 330, 760)

            # ── PAGE 1 ────────────────────────────────────────────────────
            p.setFont('Courier-Bold', 10)
            p.drawString(480, 750, draft)
            p.setFont('Courier', 10)
            p.drawString(400, 750, "PERMIT NO : ")

            p.drawString(rgtcol, 700, "CARGO CLEARANCE PERMIT")
            p.drawString(460, 700, f"PG : {countPage} OF")
            countPage += 1

            p.drawString(lftcol, 670, "MESSAGE TYPE      : IN-PAYMENT PERMIT")
            decltype = str(Inheader.get('DeclarationType', '') or '')[6:].upper()
            p.drawString(lftcol, 660, "DECLARATION TYPE  : " + decltype)

            p.drawString(lftcol, 630, "IMPORTER:")
            if len(importerName) >= 35:
                p.drawString(lftcol, 620, importerName[:35].upper())
                p.drawString(lftcol, 610, importerName[35:].upper())
            else:
                p.drawString(lftcol, 620, importerName.upper())

            p.drawString(lftcol, 600, (Impo.get('CRUEI', '') or '').upper())
            p.drawString(lftcol, 590, "EXPORTER:")
            p.drawString(lftcol, 580, " ")
            p.drawString(lftcol, 570, " ")
            p.drawString(lftcol, 560, " ")
            p.drawString(lftcol, 550, "HANDLING AGENT: ")
            p.drawString(lftcol, 540, " ")
            p.drawString(lftcol, 530, " ")
            p.drawString(lftcol, 520, " ")
            p.drawString(lftcol, 510, " ")
            p.drawString(lftcol, 500, "PORT OF LOADING/NEXT PORT OF CALL:")
            p.drawString(lftcol, 490, (loadingPort.get('portname', '') or '').upper())
            p.drawString(lftcol, 480, "PORT OF DISCHARGE/FINAL PORT OF CALL ")
            p.drawString(lftcol, 470, " ")
            p.drawString(lftcol, 460, "COUNTRY OF FINAL DESTINATION:")
            p.drawString(lftcol, 450, " ")
            p.drawString(lftcol, 440, "INWARD CARRIER AGENT: ")

            if len(InwardCarrierName) >= 35:
                p.drawString(lftcol, 430, InwardCarrierName[:35].upper())
                p.drawString(lftcol, 420, InwardCarrierName[35:70].upper())
                p.drawString(lftcol, 410, InwardCarrierName[70:].upper())
            else:
                p.drawString(lftcol, 430, InwardCarrierName.upper())

            p.drawString(lftcol, 400, "OUTWARD CARRIER AGENT: ")
            p.drawString(lftcol, 390, " ")
            p.drawString(lftcol, 380, " ")
            p.drawString(lftcol, 370, " ")
            p.drawString(lftcol, 360, "PLACE OF RELEASE: ")

            rely = 350
            relVal = (Inheader.get('ResLoaName', '') or Inheader.get('ReleaseLocName', '') or '').upper()
            chr_idx = 0
            if len(relVal) >= 32:
                for i in range(len(relVal)):
                    if (i + 1) % 32 == 0:
                        p.drawString(lftcol, rely, relVal[chr_idx:i])
                        chr_idx = i + 1
                        rely -= 10
                p.drawString(lftcol, rely, relVal[chr_idx:])
                rely -= 100
            else:
                p.drawString(lftcol, rely, relVal)
                rely -= 10

            p.drawString(lftcol, rely, Inheader.get('ReleaseLocation', '') or '')

            p.drawString(lftcol, 250, "LICENCE NO:")
            licence = ((Inheader.get('License', '') or '').upper()).split(',')
            for li_idx, li_y in enumerate([240, 230, 220, 210, 200]):
                p.drawString(lftcol, li_y, licence[li_idx] if li_idx < len(licence) else '')

            p.drawString(lftcol, 50, "-" * 80)
            p.drawString(lftcol, 40, f"UNIQUE REF : {(Declarant.get('CRUEI','') or '').upper()} {MSGId[:8].upper()} {MSGId[8:].upper()}")

            # ── Right column – page 1 ─────────────────────────────────────
            rgy = 630
            p.drawString(rgtcol, rgy, f"VALIDITY PERIOD      : {inpstrtdate} - ")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"                       {inpEndate}")
            rgy -= 20

            totalGross = "{:.3f}".format(float(Inheader.get('TotalGrossWeight') or 0))
            p.drawString(rgtcol, rgy, f"TOTAL GROSS WT/UNIT  : {totalGross:>18}/{Inheader.get('TotalGrossWeightUOM','')}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOTAL OUTER PACK/UNIT: {str(Inheader.get('TotalOuterPack','') or ''):>18}/{Inheader.get('TotalOuterPackUOM','')}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOT EXCISE DUT PAYABLE  : S${str(Inheader.get('TotalExDutyAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOT CUSTOMS DUT PAYABLE : S${str(Inheader.get('TotalCusDutyAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOT OTHER TAX PAYABLE   : S${str(Inheader.get('TotalODutyAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOTAL GST AMT           : S${str(Inheader.get('TotalGSTTaxAmt','') or ''):>17}")
            rgy -= 10
            p.drawString(rgtcol, rgy, f"TOTAL AMOUNT PAYABLE    : S${str(Inheader.get('TotalAmtPay','') or ''):>17}")
            rgy -= 10
            cargoPack = str(Inheader.get('CargoPackType', '') or '')[3:].upper()
            p.drawString(rgtcol, rgy, f"CARGO PACKING TYPE: {cargoPack} ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "IN TRANSPORT IDENTIFIER: ")
            rgy -= 10

            inwardMode = str(Inheader.get('InwardTransportMode', '') or '')[4:].upper()
            if inwardMode in ("SEA", "AIR"):
                transId = Inheader.get('VesselName', '') or ''
            else:
                transId = Inheader.get('TransportId', '') or ''

            if inwardMode == "SEA":
                CoveyanceNo = Inheader.get('VoyageNumber', '') or ''
            elif inwardMode == "AIR":
                CoveyanceNo = Inheader.get('FlightNO', '') or ''
            else:
                CoveyanceNo = Inheader.get('ConveyanceRefNo', '') or ''

            p.drawString(rgtcol, rgy, transId.upper())
            rgy -= 10
            p.drawString(rgtcol, rgy, f"CONVEYANCE REFERENCE NO: {CoveyanceNo.upper()} ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "OBL/MAWB NO: ")
            rgy -= 10
            p.drawString(rgtcol, rgy, Inheader.get('OceanBillofLadingNo', '') or '')
            rgy -= 10
            arrDate = Inheader.get('ArrivalDate')
            try:
                arrDateStr = arrDate.strftime('%d/%m/%Y') if arrDate else ''
            except Exception:
                try:
                    arrDateStr = datetime.strptime(str(arrDate), '%Y-%m-%d').strftime('%d/%m/%Y')
                except Exception:
                    arrDateStr = str(arrDate) if arrDate else ''
            p.drawString(rgtcol, rgy, f"ARRIVAL DATE         : {arrDateStr}")
            rgy -= 10
            p.drawString(rgtcol, rgy, "OU TRANSPORT IDENTIFIER: ")
            rgy -= 10
            p.drawString(rgtcol, rgy, " ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "CONVEYANCE REFERENCE NO:  ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "OBL/MAWB/UCR NO: ")
            rgy -= 10
            p.drawString(rgtcol, rgy, " ")
            rgy -= 10
            p.drawString(rgtcol, rgy, "DEPARTURE DATE       : ")
            rgy -= 20
            p.drawString(rgtcol, rgy, "CERTIFICATE NO:  ")
            rgy -= 30
            p.drawString(rgtcol, rgy, "PLACE OF RECEIPT:")
            rgy -= 10

            rely = rgy
            recVal = Inheader.get('RecepitLocName', '') or ''
            chr1 = 0
            if len(recVal) >= 32:
                for i in range(len(recVal)):
                    if (i + 1) % 32 == 0:
                        p.drawString(rgtcol, rely, recVal[chr1:i])
                        chr1 = i + 1
                        rely -= 10
                p.drawString(rgtcol, rely, recVal[chr1:])
                rely -= 10
            else:
                p.drawString(rgtcol, rely, recVal)
                rely -= 10

            p.drawString(rgtcol, rely, Inheader.get('RecepitLocation', '') or '')
            p.drawString(rgtcol, 250, "CUSTOMS PROCEDURE CODE (CPC) : ")
            rely = 240
            for cpc_row in Cpcs:
                p.drawString(rgtcol, rely, cpc_row.get('CPCType', '') or '')
                rely -= 10
            if str(Inheader.get('Cnb', '') or '').lower() == 'true':
                p.drawString(rgtcol, rely, "CNB")
                rely -= 10

            p.showPage()
            # ── PAGE 1 COMPLETE ───────────────────────────────────────────

            # ── Item page helpers ─────────────────────────────────────────
            snox    = 50
            hscodex = 100
            currentx = 180
            prviousx = 320
            makingx  = 50
            cityx    = 120
            brandx   = 220
            itemy    = 820

            def itemyF(itemy):
                nonlocal countPage
                if itemy <= 70:
                    itemy = 820
                    p.showPage()
                    p.setFont('Courier', 10)
                else:
                    itemy -= 10

                if 700 <= itemy <= 820:
                    p.setFont('Courier', 10)
                    p.drawString(rgtcol, itemy, "CARGO CLEARANCE PERMIT ")
                    p.drawString(460, itemy, f"PG : {countPage} OF ")
                    countPage += 1
                    itemy -= 10
                    p.drawString(lftcol, itemy, "PERMIT NO : " + draft)
                    p.drawString(rgtcol, itemy, "======================")
                    itemy -= 10
                    p.drawString(rgtcol, itemy, "(CONTINUATION PAGE)")
                    itemy -= 20
                    p.drawString(lftcol, itemy, "CONSIGNMENT DETAILS")
                    itemy -= 10
                    p.drawString(lftcol, itemy, "-" * 80)
                    itemy -= 10
                    p.drawString(snox, itemy, "S/NO ")
                    p.drawString(hscodex, itemy, "HS CODE")
                    p.drawString(currentx, itemy, "CURRENT LOT NO")
                    p.drawString(prviousx, itemy, "PREVIOUS LOT NO")
                    itemy -= 10
                    p.drawString(makingx, itemy, "MARKING")
                    p.drawString(cityx, itemy, "CTY OF ORIGIN")
                    p.drawString(brandx, itemy, "BRAND NAME")
                    p.drawString(prviousx, itemy, "MODEL")
                    itemy -= 10
                    # HAWB header row if any item has InHAWBOBL
                    for itd in ItemDtls:
                        if itd.get('InHAWBOBL', ''):
                            p.drawString(lftcol, itemy, "IN HAWB/HUCR/HBL")
                            p.drawString(prviousx, itemy, "OUT HAWB/HUCR/HBL")
                            itemy -= 10
                            break
                    p.drawString(lftcol, itemy, "PACKING/GOODS DESCRIPTION")
                    p.drawString(prviousx, itemy, "HS QUANTITY & UNIT ")
                    itemy -= 10
                    p.drawString(prviousx, itemy, "CIF/FOB VALUE (S$) ")
                    itemy -= 10
                    for itd in ItemDtls:
                        if float(itd.get('LSPValue') or 0) != 0.0:
                            p.drawString(prviousx, itemy, "LSP VALUE (S$) ")
                            itemy -= 10
                            break
                    p.drawString(prviousx, itemy, 'GST AMOUNT (S$)')
                    itemy -= 10
                    for itd in ItemDtls:
                        if float(itd.get('TotalDutiableQty') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'DUT QTY/WT/VOL & UNIT')
                            itemy -= 10
                            break
                    for itd in ItemDtls:
                        if float(itd.get('ExciseDutyAmount') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'EXCISE DUTY PAYABLE (S$)')
                            itemy -= 10
                            break
                    for itd in ItemDtls:
                        if float(itd.get('CustomsDutyAmount') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'CUSTOMS DUTY PAYABLE(S$)')
                            itemy -= 10
                            break
                    for itd in ItemDtls:
                        if float(itd.get('OtherTaxAmount') or 0) != 0.0:
                            p.drawString(prviousx, itemy, 'OTHER TAX PAYABLE(S$) ')
                            itemy -= 10
                            break
                    p.drawString(snox, itemy, "MANUFACTURER'S NAME ")
                    itemy -= 10
                    p.drawString(snox, itemy, '-' * 79)
                    itemy -= 10
                    p.drawString(lftcol, 50, "-" * 80)
                    p.drawString(lftcol, 40, f"UNIQUE REF : {(Declarant.get('CRUEI','') or '').upper()} {MSGId[:8].upper()} {MSGId[8:].upper()}")
                return itemy

            # ── helper: write a right-aligned numeric value ───────────────
            def write_val_right(val, y_pos, is_qty=False, uom=''):
                try:
                    num = float(val or 0)
                    dec_parts = str(val).split('.')
                    dec_part  = dec_parts[1] if len(dec_parts) > 1 else '0'
                    if is_qty:
                        p.drawString(prviousx + 100, y_pos, f'{int(num):8d}.{dec_part} {uom}')
                    else:
                        p.drawString(prviousx + 100, y_pos, f'{int(num):10d}.{dec_part}')
                except Exception:
                    pass

            # ── ITEMS LOOP ────────────────────────────────────────────────
            for item in ItemDtls:
                itemy = itemyF(itemy)
                p.drawString(snox, itemy, f"   {int(item.get('ItemNo', 0)):02d}")
                p.drawString(hscodex, itemy, str(item.get('HSCode', '') or ''))
                p.drawString(currentx, itemy, (item.get('CurrentLot', '') or '').upper())
                p.drawString(prviousx, itemy, (item.get('PreviousLot', '') or '').upper())

                itemy = itemyF(itemy)
                making = str(item.get('Making', '') or '').upper()
                if making not in ('--SELECT--', 'NONE', ''):
                    p.drawString(makingx, itemy, making[:2])
                p.drawString(cityx - 40, itemy, (item.get('Contry', '') or '').upper())
                p.drawString(brandx - 105, itemy, (item.get('Brand', '') or '').upper())
                p.drawString(prviousx, itemy, (item.get('Model', '') or '').upper())

                itemy = itemyF(itemy)
                if item.get('InHAWBOBL', ''):
                    p.drawString(lftcol, itemy, (item.get('InHAWBOBL', '') or '').upper())
                    itemy = itemyF(itemy)

                ItemDescr = str(item.get('Description', '') or '').upper()

                if '\n' in ItemDescr:
                    itemDesc = ItemDescr.split('\n')

                    def get_desc(i):
                        return itemDesc[i].replace('\r', '') if i < len(itemDesc) else ''

                    # Line 0 – HS Qty
                    p.drawString(lftcol, itemy, get_desc(0))
                    write_val_right(item.get('HSQty'), itemy, True, item.get('HSUOM', ''))
                    itemy = itemyF(itemy)

                    # Line 1 – CIF/FOB
                    p.drawString(lftcol, itemy, get_desc(1))
                    write_val_right(item.get('CIFFOB'), itemy)
                    itemy = itemyF(itemy)

                    # Line 2 – LSP (optional)
                    if float(item.get('LSPValue') or 0) != 0.0:
                        p.drawString(lftcol, itemy, get_desc(2))
                        write_val_right(item.get('LSPValue'), itemy)
                        itemy = itemyF(itemy)
                    else:
                        p.drawString(lftcol, itemy, get_desc(2))

                    # GST Amount
                    write_val_right(item.get('GSTAmount'), itemy)
                    itemy = itemyF(itemy)

                    # Line 3 – Dutiable Qty (optional)
                    if float(item.get('TotalDutiableQty') or 0) != 0.0:
                        p.drawString(lftcol, itemy, get_desc(3))
                        write_val_right(item.get('TotalDutiableQty'), itemy, True, item.get('TotalDutiableUOM', ''))
                        itemy = itemyF(itemy)
                    elif len(itemDesc) >= 4:
                        p.drawString(lftcol, itemy, get_desc(3))
                        itemy = itemyF(itemy)

                    # Lines 4-6 – Duty amounts
                    for duty_field, desc_idx in [
                        ('ExciseDutyAmount', 4),
                        ('CustomsDutyAmount', 5),
                        ('OtherTaxAmount', 6),
                    ]:
                        if float(item.get(duty_field) or 0) != 0.0:
                            p.drawString(lftcol, itemy, get_desc(desc_idx))
                            write_val_right(item.get(duty_field), itemy)
                            itemy = itemyF(itemy)
                        elif len(itemDesc) >= desc_idx + 1:
                            p.drawString(lftcol, itemy, get_desc(desc_idx))
                            itemy = itemyF(itemy)

                else:
                    # ── Flat description (no newlines) ────────────────────
                    slices = [
                        (0,   50),
                        (50,  100),
                        (100, 150),
                        (150, 200),
                        (200, 250),
                        (250, 300),
                        (300, 350),
                    ]
                    fields = [
                        ('HSQty',            True,  item.get('HSUOM', '')),
                        ('CIFFOB',           False, ''),
                        ('LSPValue',         False, ''),
                        ('GSTAmount',        False, ''),
                        ('TotalDutiableQty', True,  item.get('TotalDutiableUOM', '')),
                        ('ExciseDutyAmount', False, ''),
                        ('CustomsDutyAmount',False, ''),
                        ('OtherTaxAmount',   False, ''),
                    ]
                    for idx, (s, e) in enumerate(slices):
                        if len(ItemDescr) >= s or idx == 0:
                            p.drawString(lftcol, itemy, ItemDescr[s:e])
                            if idx < len(fields):
                                fname, is_qty, uom = fields[idx]
                                val = item.get(fname)
                                # skip LSP if zero
                                if fname == 'LSPValue' and float(val or 0) == 0.0:
                                    pass
                                elif val is not None:
                                    write_val_right(val, itemy, is_qty, uom)
                            itemy = itemyF(itemy)

                    # Extra description chunks beyond 350
                    for s in range(350, 600, 50):
                        chunk = ItemDescr[s:s + 50]
                        if len(ItemDescr) >= s and chunk:
                            p.drawString(lftcol, itemy, chunk)
                            itemy = itemyF(itemy)

                # ── Manufacturer name from matching invoice ────────────────
                for inv in InvoiceDtls:
                    if inv.get('InvoiceNo') == item.get('InvoiceNo'):
                        manfs = SqlDb.execute_query(
                            "SELECT * FROM CommonSupplierManufacturerPart WHERE Code = %s",
                            [inv.get('SupplierCode')]
                        )
                        if manfs:
                            manf     = manfs[0]
                            manfName = manf.get('Name', '') or ''
                            p.drawString(lftcol, itemy, manfName[:50])
                            itemy = itemyF(itemy)
                            if len(manfName) >= 50:
                                p.drawString(lftcol, itemy, manfName[50:])
                        break

                # ── CASC details ──────────────────────────────────────────
                Cascdtls = SqlDb.execute_query(
                    "SELECT * FROM CommonCASCDtl WHERE PermitId = %s AND ItemNo = %s",
                    [permit_id, item.get('ItemNo')]
                )
                if Cascdtls:
                    for casc_sno, casc_id_key in [
                        ('01', 'Casc1'), ('02', 'Casc2'), ('03', 'Casc3'),
                        ('04', 'Casc4'), ('05', 'Casc5'),
                    ]:
                        matched = [c for c in Cascdtls if c.get('CASCId') == casc_id_key]
                        if matched:
                            casc = matched[0]
                            p.drawString(lftcol, itemy, "-" * 80)
                            itemy = itemyF(itemy)
                            p.drawString(lftcol, itemy, 'S/NO')
                            p.drawString(lftcol + 70, itemy, 'CA/SC PRODUCT CODE ')
                            p.drawString(lftcol + 280, itemy, 'CA/SC PRODUCT QTY & UNIT')
                            itemy = itemyF(itemy)
                            p.drawString(lftcol, itemy, f'   {casc_sno}')
                            p.drawString(lftcol + 70, itemy, (casc.get('ProductCode', '') or '').upper())
                            qty = casc.get('Quantity')
                            if qty and str(qty) not in ("0.0000", "0.00", "0"):
                                try:
                                    dec = str(qty).split('.')
                                    p.drawString(lftcol + 280, itemy,
                                        f"{int(float(qty)):10d}.{dec[1] if len(dec)>1 else '0'} {casc.get('ProductUOM','')}")
                                except Exception:
                                    pass
                            itemy = itemyF(itemy)
                            p.drawString(lftcol, itemy, "-" * 80)
                            break

                # ── Engine capacity ───────────────────────────────────────
                engCap = str(item.get('EngineCapcity', '') or '')
                if engCap not in ("0.00", "0", ""):
                    p.drawString(snox, itemy, '-' * 79)
                    itemy = itemyF(itemy)
                    p.drawString(lftcol, itemy, 'S/NO')
                    p.drawString(lftcol + 70, itemy, 'ENGINE NO/CHASSIS NO ')
                    itemy = itemyF(itemy)
                    p.drawString(lftcol, itemy, f"   {int(item.get('ItemNo', 0)):02d}")
                    p.drawString(lftcol + 70, itemy, engCap)
                    itemy = itemyF(itemy)
                    p.drawString(snox, itemy, '-' * 79)

                p.drawString(snox, itemy, '-' * 79)
                itemy = itemyF(itemy) - 20
                p.drawString(snox, itemy, '-' * 79)

            # ── Trailing sections helper ───────────────────────────────────
            def itemyF1(itemy):
                nonlocal countPage
                if itemy <= 60:
                    itemy = 820
                    p.showPage()
                    p.setFont('Courier', 10)
                else:
                    itemy -= 10
                if 780 <= itemy <= 820:
                    p.setFont('Courier', 10)
                    p.drawString(rgtcol, itemy, "CARGO CLEARANCE PERMIT ")
                    p.drawString(460, itemy, f"PG : {countPage} OF")
                    countPage += 1
                    itemy -= 10
                    p.drawString(lftcol, itemy, "PERMIT NO : " + draft)
                    p.drawString(rgtcol, itemy, "======================")
                    itemy -= 10
                    p.drawString(rgtcol, itemy, "(CONTINUATION PAGE)")
                    itemy -= 20
                    p.drawString(lftcol, itemy, "CONSIGNMENT DETAILS")
                    itemy -= 10
                    p.drawString(lftcol, itemy, "-" * 80)
                    itemy -= 10
                    p.drawString(lftcol, 50, "-" * 80)
                    p.drawString(lftcol, 40,
                        f"UNIQUE REF : {(Declarant.get('CRUEI','') or '').upper()} {MSGId[:8].upper()} {MSGId[8:].upper()}")
                return itemy

            # ── Trader's remarks ──────────────────────────────────────────
            tradeRemarks = str(Inheader.get('TradeRemarks', '') or '')
            if tradeRemarks:
                itemy = itemyF1(itemy)
                p.drawString(lftcol, itemy, "TRADER'S REMARKS")
                itemy = itemyF1(itemy)
                TradeRe = tradeRemarks.upper()
                if '\n' in TradeRe:
                    for trd in TradeRe.split('\n'):
                        p.drawString(lftcol, itemy, trd)
                        itemy = itemyF1(itemy)
                else:
                    for s in range(0, max(len(TradeRe), 1), 80):
                        chunk = TradeRe[s:s + 80]
                        if chunk:
                            p.drawString(lftcol, itemy, chunk)
                            itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, '-' * 79)

            # ── Container identifiers ─────────────────────────────────────
            if ContainerDtls:
                itemy = itemyF1(itemy)
                p.drawString(lftcol, itemy, "CONTAINER IDENTIFIERS")
                itemy = itemyF1(itemy)
                for cont in ContainerDtls:
                    p.drawString(lftcol, itemy, f"   {int(cont.get('RowNo', 0)):02d})")
                    p.drawString(lftcol + 50, itemy, str(cont.get('ContainerNo', '') or '').upper())
                    size = str(cont.get('Size', '') or '')
                    p.drawString(lftcol + 130, itemy, f"{size[:3]}  {size[3:5]}")
                    p.drawString(lftcol + 200, itemy, str(cont.get('Weight', '') or '')[:3])
                    p.drawString(lftcol + 240, itemy, str(cont.get('SealNo', '') or ''))
                    itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, '-' * 79)
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, "NO UNAUTHORISED ADDITION/AMENDMENT TO THIS PERMIT MAY BE MADE AFTER APPROVAL")
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, '-' * 79)
            itemy = itemyF1(itemy)

            # ── Declarant block ───────────────────────────────────────────
            p.drawString(lftcol, itemy, "NAME OF COMPANY:")
            decName = Declarant.get('name', '') or ''
            p.drawString(lftcol + 110, itemy, decName[:67])
            itemy = itemyF1(itemy)
            p.drawString(lftcol + 110, itemy, decName[67:])
            itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, "DECLARANT NAME :")
            decPersonName = Declarant.get('DeclarantName', '') or ''
            p.drawString(lftcol + 110, itemy, decPersonName[:67])
            itemy = itemyF1(itemy)
            p.drawString(lftcol + 110, itemy, decPersonName[67:])
            itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, "DECLARANT CODE :")
            decCode = Declarant.get('DeclarantCode', '') or ''
            p.drawString(lftcol + 110, itemy, "XXXX" + decCode[-5:])
            itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, "TEL NO         : ")
            p.drawString(lftcol + 110, itemy, Declarant.get('DeclarantTel', '') or '')
            itemy = itemyF1(itemy)
            p.drawString(lftcol, itemy, '-' * 79)
            itemy = itemyF1(itemy)

            p.drawString(lftcol, itemy, "CONTROLLING AGENCY/CUSTOMS CONDITIONS ")
            itemy = itemyF1(itemy)

            for pmt in pmts:
                p.setFont('Courier-Bold', 10)
                p.drawString(lftcol, itemy, pmt.get('ConditionCode', '') or '')
                p.drawString(lftcol + 30, itemy, "-")
                p.setFont('Courier', 10)
                condDesc = pmt.get('ConditionDesc', '') or pmt.get('ConditionDescription', '') or ''
                p.drawString(lftcol + 40, itemy, condDesc[:73])
                itemy = itemyF1(itemy)
                for s in range(73, len(condDesc), 80):
                    p.drawString(lftcol, itemy, condDesc[s:s + 80])
                    itemy = itemyF1(itemy)

            # ── Save first-pass PDF to temp buffer ────────────────────────
            p.setTitle(draft)
            p.save()

            # ── Overlay total page count (merge pass) ─────────────────────
            buffer.seek(0)
            existing_pdf = PdfReader(buffer)
            total_pages  = len(existing_pdf.pages)
            output       = PdfWriter()

            packet2 = io.BytesIO()
            can2    = canvas.Canvas(packet2, pagesize=(595, 841))
            can2.setFont('Courier', 10)
            can2.drawString(520, 700, str(total_pages))   # page 1 overlay
            can2.showPage()
            can2.setFont('Courier', 10)
            can2.drawString(520, 810, str(total_pages))   # page 2 overlay
            can2.showPage()
            can2.setFont('Courier', 10)
            can2.drawString(520, 820, str(total_pages))   # page 3+ overlay
            can2.save()
            packet2.seek(0)
            new_pdf = PdfReader(packet2)

            for i in range(total_pages):
                page = existing_pdf.pages[i]
                if i == 0:
                    page.merge_page(new_pdf.pages[0])
                elif i == 1:
                    page.merge_page(new_pdf.pages[1])
                else:
                    page.merge_page(new_pdf.pages[2])
                output.add_page(page)

            final_buffer = io.BytesIO()
            output.write(final_buffer)
            final_buffer.seek(0)

            safe_name = draft if draft != "DRAFT" else f"Draft_{MSGId.strip()}"
            response = HttpResponse(final_buffer, content_type='application/pdf')
            response['Content-Disposition'] = f'attachment; filename="{safe_name}.pdf"'
            return response

        except Exception as e:
            import traceback
            traceback.print_exc()
            return HttpResponse(f"Error generating CCP: {str(e)}", status=500)


class DownloadData(APIView):
    def get(self, request):
        data = request.GET.get("data", "")
        if not data:
            return HttpResponse("No permits provided", status=400)

        # try:
        #     import xlwt
        # except ImportError:
        #     return HttpResponse("xlwt not installed. Run: pip install xlwt", status=500)

        wb = xlwt.Workbook(encoding='utf-8')

        # ── Styles ────────────────────────────────────────────────────
        header_style = xlwt.XFStyle()
        header_style.font.bold = True
        bg = xlwt.Pattern()
        bg.pattern = xlwt.Pattern.SOLID_PATTERN
        bg.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']
        header_style.pattern = bg

        odd_style  = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue;')
        even_style = xlwt.easyxf('pattern: pattern solid, fore_colour white;')
        data_style = xlwt.XFStyle()
        data_style.font.bold = False

        # ── Sheet definitions ─────────────────────────────────────────
        header_cols = [
            'Id', 'Refid', 'JobId', 'MSGId', 'PermitId', 'TradeNetMailboxID',
            'MessageType', 'DeclarationType', 'PreviousPermit', 'CargoPackType',
            'InwardTransportMode', 'BGIndicator', 'SupplyIndicator',
            'ReferenceDocuments', 'License', 'Recipient', 'DeclarantCompanyCode',
            'ImporterCompanyCode', 'InwardCarrierAgentCode', 'FreightForwarderCode',
            'ClaimantPartyCode', 'HBL', 'ArrivalDate', 'LoadingPortCode',
            'VoyageNumber', 'VesselName', 'OceanBillofLadingNo', 'ConveyanceRefNo',
            'TransportId', 'FlightNO', 'AircraftRegNo', 'MasterAirwayBill',
            'ReleaseLocation', 'ReleaseLocName', 'RecepitLocation', 'TotalOuterPack',
            'TotalOuterPackUOM', 'TotalGrossWeight', 'TotalGrossWeightUOM',
            'GrossReference', 'BlanketStartDate', 'TradeRemarks', 'InternalRemarks',
            'CustomerRemarks', 'DeclareIndicator', 'NumberOfItems',
            'TotalCIFFOBValue', 'TotalGSTTaxAmt', 'TotalExDutyAmt',
            'TotalCusDutyAmt', 'TotalODutyAmt', 'TotalAmtPay', 'Status',
            'TouchUser', 'TouchTime', 'PermitNumber', 'prmtStatus',
            'RecepitLocName', 'Cnb', 'DeclarningFor', 'MRDate', 'MRTime',
        ]
        # DB columns that map to header_cols (Id comes from auto primary key)
        header_db_cols = [
            'Id', 'Refid', 'JobId', 'MSGId', 'PermitId', 'TradeNetMailboxID',
            'MessageType', 'DeclarationType', 'PreviousPermit', 'CargoPackType',
            'InwardTransportMode', 'BGIndicator', 'SupplyIndicator',
            'ReferenceDocuments', 'License', 'Recipient', 'DeclarantCompanyCode',
            'ImporterCompanyCode', 'InwardCarrierAgentCode', 'FreightForwarderCode',
            'ClaimantPartyCode', 'HBL', 'ArrivalDate', 'LoadingPortCode',
            'VoyageNumber', 'VesselName', 'OceanBillofLadingNo', 'ConveyanceRefNo',
            'TransportId', 'FlightNO', 'AircraftRegNo', 'MasterAirwayBill',
            'ReleaseLocation', 'ResLoaName', 'RecepitLocation', 'TotalOuterPack',
            'TotalOuterPackUOM', 'TotalGrossWeight', 'TotalGrossWeightUOM',
            'GrossReference', 'BlanketStartDate', 'TradeRemarks', 'InternalRemarks',
            'InternalRemarks', 'DeclareIndicator', 'NumberOfItems',
            'TotalCIFFOBValue', 'TotalGSTTaxAmt', 'TotalExDutyAmt',
            'TotalCusDutyAmt', 'TotalODutyAmt', 'TotalAmtPay', 'Status',
            'TouchUser', 'TouchTime', 'PermitNumber', 'prmtStatus',
            'RecepitLocName', 'Cnb', 'DeclarningFor', 'MRDate', 'MRTime',
        ]

        container_cols = [
            'Id', 'PermitId', 'RowNo', 'ContainerNo', 'Size', 'Weight',
            'SealNo', 'MessageType', 'TouchUser', 'TouchTime',
        ]

        invoice_cols = [
            'Id', 'SNo', 'InvoiceNo', 'InvoiceDate', 'TermType',
            'AdValoremIndicator', 'PreDutyRateIndicator',
            'SupplierImporterRelationship', 'SupplierCode', 'ImportPartyCode',
            'TICurrency', 'TIExRate', 'TIAmount', 'TISAmount',
            'OTCCharge', 'OTCCurrency', 'OTCExRate', 'OTCAmount', 'OTCSAmount',
            'FCCharge', 'FCCurrency', 'FCExRate', 'FCAmount', 'FCSAmount',
            'ICCharge', 'ICCurrency', 'ICExRate', 'ICAmount', 'ICSAmount',
            'CIFSUMAmount', 'GSTPercentage', 'GSTSUMAmount',
            'MessageType', 'PermitId', 'TouchUser', 'TouchTime', 'ChkOtherInv',
        ]

        item_cols = [
            'Id', 'ItemNo', 'PermitId', 'MessageType', 'HSCode', 'Description',
            'DGIndicator', 'Contry', 'Brand', 'Model', 'InHAWBOBL',
            'DutiableQty', 'DutiableUOM', 'TotalDutiableQty', 'TotalDutiableUOM',
            'HSQty', 'HSUOM', 'AlcoholPer', 'InvoiceNo', 'ChkUnitPrice',
            'UnitPrice', 'UnitPriceCurrency', 'ExchangeRate', 'SumExchangeRate',
            'TotalLineAmount', 'InvoiceCharges', 'CIFFOB',
            'OPQty', 'OPUOM', 'IPQty', 'IPUOM', 'InPqty', 'InPUOM',
            'ImPQty', 'ImPUOM', 'PreferentialCode', 'GSTRate', 'GSTUOM',
            'GSTAmount', 'ExciseDutyRate', 'ExciseDutyUOM', 'ExciseDutyAmount',
            'CustomsDutyRate', 'CustomsDutyUOM', 'CustomsDutyAmount',
            'OtherTaxRate', 'OtherTaxUOM', 'OtherTaxAmount',
            'CurrentLot', 'PreviousLot', 'LSPValue', 'Making',
            'ShippingMarks1', 'ShippingMarks2', 'ShippingMarks3', 'ShippingMarks4',
            'TouchUser', 'TouchTime', 'VehicleType', 'EngineCapcity',
            'EngineCapUOM', 'orignaldatereg', 'OptionalChrgeUOM',
            'Optioncahrge', 'OptionalSumtotal', 'OptionalSumExchage',
            'InvoiceQuantity',
        ]

        casc_cols = [
            'Id', 'ItemNo', 'ProductCode', 'Quantity', 'ProductUOM', 'RowNo',
            'CascCode1', 'CascCode2', 'CascCode3', 'PermitId', 'MessageType',
            'TouchUser', 'TouchTime', 'CASCId',
        ]

        # ── Create sheets & write header rows ─────────────────────────
        header_sheet    = wb.add_sheet('Header')
        container_sheet = wb.add_sheet('Container')
        invoice_sheet   = wb.add_sheet('Invoice')
        item_sheet      = wb.add_sheet('Item')
        casc_sheet      = wb.add_sheet('Casc')

        for ci, col in enumerate(header_cols):
            header_sheet.write(0, ci, col, header_style)
        for ci, col in enumerate(container_cols):
            container_sheet.write(0, ci, col, header_style)
        for ci, col in enumerate(invoice_cols):
            invoice_sheet.write(0, ci, col, header_style)
        for ci, col in enumerate(item_cols):
            item_sheet.write(0, ci, col, header_style)
        for ci, col in enumerate(casc_cols):
            casc_sheet.write(0, ci, col, header_style)

        # ── Row counters (0 = header already written) ─────────────────
        h_row = c_row = inv_row = itm_row = casc_row = 0

        def write_row(sheet, row_num, values):
            style = odd_style if row_num % 2 == 0 else even_style
            for ci, val in enumerate(values):
                # xlwt can't handle None or datetime natively for all cases
                if val is None:
                    val = ""
                try:
                    sheet.write(row_num, ci, str(val) if not isinstance(val, (int, float)) else val, style)
                except Exception:
                    sheet.write(row_num, ci, str(val), style)

        # ── Loop permits ──────────────────────────────────────────────
        for permit_id in data.split(','):
            permit_id = permit_id.strip()
            if not permit_id:
                continue

            # ── Header ────────────────────────────────────────────────
            rows = SqlDb.execute_query(
                f"""
                SELECT
                    Id, Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                    MessageType, DeclarationType, PreviousPermit, CargoPackType,
                    InwardTransportMode, BGIndicator, SupplyIndicator,
                    ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                    ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                    ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
                    VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                    TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                    ReleaseLocation, ResLoaName, RecepitLocation, TotalOuterPack,
                    TotalOuterPackUOM, TotalGrossWeight, TotalGrossWeightUOM,
                    GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                    InternalRemarks, DeclareIndicator, NumberOfItems,
                    TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                    TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay, Status,
                    TouchUser, TouchTime, PermitNumber, prmtStatus,
                    RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime
                FROM CommonHeaderTbl
                WHERE PermitId = %s
                """,
                [permit_id]
            )
            for row in rows:
                h_row += 1
                write_row(header_sheet, h_row, list(row.values()))

            # ── Container ─────────────────────────────────────────────
            rows = SqlDb.execute_query(
                """
                SELECT Id, PermitId, RowNo, ContainerNo, Size, Weight,
                       SealNo, MessageType, TouchUser, TouchTime
                FROM CommonContainerDtl
                WHERE PermitId = %s
                ORDER BY RowNo
                """,
                [permit_id]
            )
            for row in rows:
                c_row += 1
                write_row(container_sheet, c_row, list(row.values()))

            # ── Invoice ───────────────────────────────────────────────
            rows = SqlDb.execute_query(
                """
                SELECT Id, SNo, InvoiceNo, InvoiceDate, TermType,
                       AdValoremIndicator, PreDutyRateIndicator,
                       SupplierImporterRelationship, SupplierCode, ImportPartyCode,
                       TICurrency, TIExRate, TIAmount, TISAmount,
                       OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                       FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                       ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                       CIFSUMAmount, GSTPercentage, GSTSUMAmount,
                       MessageType, PermitId, TouchUser, TouchTime, ChkOtherInv
                FROM CommonInvoiceDtl
                WHERE PermitId = %s
                ORDER BY SNo
                """,
                [permit_id]
            )
            for row in rows:
                inv_row += 1
                write_row(invoice_sheet, inv_row, list(row.values()))

            # ── Item ──────────────────────────────────────────────────
            rows = SqlDb.execute_query(
                """
                SELECT Id, ItemNo, PermitId, MessageType, HSCode, Description,
                       DGIndicator, Contry, Brand, Model, InHAWBOBL,
                       DutiableQty, DutiableUOM, TotalDutiableQty, TotalDutiableUOM,
                       HSQty, HSUOM, AlcoholPer, InvoiceNo, ChkUnitPrice,
                       UnitPrice, UnitPriceCurrency, ExchangeRate, SumExchangeRate,
                       TotalLineAmount, InvoiceCharges, CIFFOB,
                       OPQty, OPUOM, IPQty, IPUOM, InPqty, InPUOM, ImPQty, ImPUOM,
                       PreferentialCode, GSTRate, GSTUOM, GSTAmount,
                       ExciseDutyRate, ExciseDutyUOM, ExciseDutyAmount,
                       CustomsDutyRate, CustomsDutyUOM, CustomsDutyAmount,
                       OtherTaxRate, OtherTaxUOM, OtherTaxAmount,
                       CurrentLot, PreviousLot, LSPValue, Making,
                       ShippingMarks1, ShippingMarks2, ShippingMarks3, ShippingMarks4,
                       TouchUser, TouchTime, VehicleType, EngineCapcity,
                       EngineCapUOM, orignaldatereg, OptionalChrgeUOM,
                       Optioncahrge, OptionalSumtotal, OptionalSumExchage,
                       InvoiceQuantity
                FROM CommonItemDtl
                WHERE PermitId = %s
                ORDER BY ItemNo
                """,
                [permit_id]
            )
            for row in rows:
                itm_row += 1
                write_row(item_sheet, itm_row, list(row.values()))

            # ── Casc ──────────────────────────────────────────────────
            rows = SqlDb.execute_query(
                """
                SELECT Id, ItemNo, ProductCode, Quantity, ProductUOM, RowNo,
                       CascCode1, CascCode2, CascCode3, PermitId, MessageType,
                       TouchUser, TouchTime, CASCId
                FROM CommonCASCDtl
                WHERE PermitId = %s
                ORDER BY ItemNo
                """,
                [permit_id]
            )
            for row in rows:
                casc_row += 1
                write_row(casc_sheet, casc_row, list(row.values()))

        # ── Return XLS response ───────────────────────────────────────
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DownloadData.xls"'
        wb.save(response)
        return response


class GstExcel(APIView):

    def get(self, request):
        data = request.GET.get("data", "")
        if not data:
            return HttpResponse("No permits provided", status=400)

        permit_ids = [p.strip() for p in data.split(",") if p.strip()]
        if not permit_ids:
            return HttpResponse("No valid permit IDs", status=400)

        xls_files = []  # list of (filename, html_content)

        for permit_id in permit_ids:
            try:
                html = self._build_gst_html(permit_id)
                if html:
                    xls_files.append(html)
            except Exception as e:
                import traceback
                traceback.print_exc()
                continue

        if not xls_files:
            return HttpResponse("No data found for given permits", status=404)

        # ── Single permit → direct XLS download ───────────────────────────
        if len(xls_files) == 1:
            filename, content = xls_files[0]
            response = HttpResponse(content, content_type="application/vnd.ms-excel")
            response["Content-Disposition"] = f'attachment; filename="{filename}"'
            return response

        # ── Multiple permits → ZIP of XLS files ───────────────────────────
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for filename, content in xls_files:
                zf.writestr(f"GST_Files/{filename}", content)

        zip_buffer.seek(0)
        response = HttpResponse(zip_buffer, content_type="application/zip")
        response["Content-Disposition"] = (
            f'attachment; filename="GST_Excel_{datetime.now().strftime("%Y%m%d_%H%M")}.zip"'
        )
        return response

    # ─────────────────────────────────────────────────────────────────────────
    def _safe(self, val):
        if val is None:
            return ""
        s = str(val).strip()
        return "" if s in ("--Select--", "None") else s

    def _fmt_num(self, val):
        """Format a number to 2 decimal places, or blank if zero/empty."""
        try:
            f = float(val or 0)
            return f"{f:.2f}" if f != 0 else "0.00"
        except Exception:
            return ""

    # ─────────────────────────────────────────────────────────────────────────
    def _build_gst_html(self, permit_id):
        """
        Build the HTML content that Excel renders as a spreadsheet.
        Returns (filename, html_string) or None if no data.
        """
        # ── Fetch header ───────────────────────────────────────────────────
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT JobId, MSGId, PermitNumber, TouchTime FROM CommonHeaderTbl WHERE PermitId = %s",
                [permit_id],
            )
            header_row = cursor.fetchone()

        if not header_row:
            return None

        job_id, msg_id, permit_number, touch_time = header_row

        # ── Fetch invoices ─────────────────────────────────────────────────
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT
                    InvoiceNo, TermType, ImportPartyCode,
                    TICurrency, TIExRate, TIAmount, TISAmount,
                    FCCharge, FCCurrency, FCExRate, FCAmount, FCSAmount,
                    ICCharge, ICCurrency, ICExRate, ICAmount, ICSAmount,
                    OTCCharge, OTCCurrency, OTCExRate, OTCAmount, OTCSAmount,
                    CIFSUMAmount, GSTPercentage, GSTSUMAmount
                FROM CommonInvoiceDtl
                WHERE PermitId = %s
                ORDER BY SNo
                """,
                [permit_id],
            )
            invoices = cursor.fetchall()

        if not invoices:
            return None

        # ── Fetch importer from first invoice ─────────────────────────────
        import_party_code = invoices[0][2]
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT Name, CRUEI FROM CommonImporter WHERE Code = %s",
                [import_party_code],
            )
            imp_row = cursor.fetchone()

        import_name = imp_row[0] if imp_row else ""
        import_uen  = imp_row[1] if imp_row else ""

        # Format date/time
        if touch_time:
            try:
                touch_time_str = touch_time.strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                touch_time_str = str(touch_time)
        else:
            touch_time_str = ""

        permit_no_display = self._safe(permit_number) or ""

        # ── Build HTML ─────────────────────────────────────────────────────
        DIVIDER = "-" * 200  # horizontal rule between invoices

        html_parts = []

        # Header section (matches template exactly)
        html_parts.append(f"""<font size='1'><table width='100%' cellspacing='0' cellpadding='2'>
            <tr><td align='center' style='background-color: #18B5F0' colspan='2'><b>GOODS AND SERVICE TAX (GST) - CALCULATION SHEET</b></td></tr>
            <tr><td colspan='2'></td></tr>
            <tr>
            <td><b>COMPANY NAME: </b>{self._safe(import_name)} </td>
            <td align='right'><b>JOB NUMBER:  </b>{self._safe(job_id)} </td>
            </tr>
            <tr>
            <td><b>COMPANY UEN: </b>{self._safe(import_uen)}</td>
            <td align='right'><b>JOB CREATED: </b>{touch_time_str} </td>
            </tr>
            <tr>
            <td><b>PERMIT NUMBER: </b>{permit_no_display}</td>
            <td align='right'><b>MESSAGE ID:    </b>{self._safe(msg_id)} </td>
            </tr>
            </table>""")

        # Invoice sections
        for inv in invoices:
            (
                invoice_no, term_type, imp_party_code,
                ti_currency, ti_ex_rate, ti_amount, ti_s_amount,
                fc_charge, fc_currency, fc_ex_rate, fc_amount, fc_s_amount,
                ic_charge, ic_currency, ic_ex_rate, ic_amount, ic_s_amount,
                otc_charge, otc_currency, otc_ex_rate, otc_amount, otc_s_amount,
                cif_sum, gst_pct, gst_sum,
            ) = inv

            html_parts.append(f"""{DIVIDER}
                <table width='100%' cellspacing='0' cellpadding='2'>
                <tr><td colspan='2'></td></tr>
                <tr>
                <td><b>INVOICE NUMBER: </b>{self._safe(invoice_no)}</td>
                <td><b>  </b> </td>
                </tr>
                <tr>
                <td><b>INVOICE TERM: </b>{self._safe(term_type)}</td>
                <td><b></b> </td>
                </tr>
                </table><br><br>
                <table width='100%' border='1'>
                <tr align='center'>
                <th>SNO</th><th>ITEM</th><th>CURRENCY</th><th>EXCHG. RATE</th>
                <th>PERC.</th><th>AMOUNT</th><th>AMOUNT (SGD)</th>
                </tr>
                <tr align='center'>
                <td>1</td><td>TOTAL INVOICE</td>
                <td>{self._safe(ti_currency)}</td><td>{self._fmt_num(ti_ex_rate)}</td>
                <td></td><td>{self._fmt_num(ti_amount)}</td><td>{self._fmt_num(ti_s_amount)}</td>
                </tr>
                <tr align='center'>
                <td>2</td><td>FREIGHT CHARGES</td>
                <td>{self._safe(fc_currency)}</td><td>{self._fmt_num(fc_ex_rate)}</td>
                <td>{self._fmt_num(fc_charge)}</td><td>{self._fmt_num(fc_amount)}</td><td>{self._fmt_num(fc_s_amount)}</td>
                </tr>
                <tr align='center'>
                <td>3</td><td>INSURANCE</td>
                <td>{self._safe(ic_currency)}</td><td>{self._fmt_num(ic_ex_rate)}</td>
                <td>{self._fmt_num(ic_charge)}</td><td>{self._fmt_num(ic_amount)}</td><td>{self._fmt_num(ic_s_amount)}</td>
                </tr>
                <tr align='center'>
                <td>4</td><td>OTHER CHARGES</td>
                <td>{self._safe(otc_currency)}</td><td>{self._fmt_num(otc_ex_rate)}</td>
                <td>{self._fmt_num(otc_charge)}</td><td>{self._fmt_num(otc_amount)}</td><td>{self._fmt_num(otc_s_amount)}</td>
                </tr>
                <tr align='center'>
                <td>5</td><td>CUSTOMS VALUE</td>
                <td></td><td></td><td></td><td></td><td>{self._fmt_num(cif_sum)}</td>
                </tr>
                <tr align='center'>
                <td>6</td><td>GST</td>
                <td></td><td></td><td>{self._safe(gst_pct)}</td><td></td><td>{self._fmt_num(gst_sum)}</td>
                </tr>
                </table>""")

        html_parts.append("</font>")

        html_content = "".join(html_parts)

        # File name: use permit number or permit_id
        safe_name = (permit_no_display or permit_id).strip().replace(" ", "_")
        filename = f"{safe_name}_GST.xls"

        return filename, html_content

# Print status
class PrintStatus(APIView):
    def get(self, request, PermitId):
        try:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(595, 841))
 
            can.setFont('Times-Roman', 9)
            can.drawString(480, 820, "PG : 1 OF 1")
 
            # ── Fetch Header ──────────────────────────────────────────────
            headers = SqlDb.execute_query(
                "SELECT * FROM CommonHeaderTbl WHERE PermitId = %s", [PermitId]
            )
            if not headers:
                return HttpResponse("Permit not found", status=404)
            Inheader = headers[0]
 
            # ── Fetch Importer ────────────────────────────────────────────
            importers = SqlDb.execute_query(
                "SELECT * FROM CommonImporter WHERE Code = %s",
                [Inheader.get('ImporterCompanyCode', '')]
            )
            Import = importers[0] if importers else {}
 
            # ── Fetch Loading Port ────────────────────────────────────────
            loadingPorts = SqlDb.execute_query(
                "SELECT * FROM LoadingPort WHERE portcode = %s",
                [Inheader.get('LoadingPortCode', '')]
            )
            Loading = loadingPorts[0] if loadingPorts else {}
 
            # ── Paragraph style ───────────────────────────────────────────
            style = getSampleStyleSheet()['Normal']
            style.fontName = 'Times-Roman'
            style.fontSize = 9
            style.spaceBefore = 0
 
            # ── Build reject/status table data ────────────────────────────
            rejectStatus = [['SERIAL NO', 'CODE', 'MESSAGE']]
            TitleStatus = ''
            Status = Inheader.get('Status', '')
            MSGId  = Inheader.get('MSGId', '')
            MailId = Inheader.get('TradeNetMailboxID', '')
 
            if Status == "REJ":
                y_ax = 430
                filename = "RejectStatus"
                TitleStatus = 'REJECTION MESSAGE'
                rows = SqlDb.execute_query(
                    """SELECT * FROM CommonRejectStatus
                       WHERE MsgId = %s AND MailBoxId = %s
                       ORDER BY Sno""",
                    [MSGId, MailId]
                )
                for re in rows:
                    rejectStatus.append([
                        Paragraph(str(re.get('Sno', '')), style),
                        Paragraph(str(re.get('ErrorId', '')), style),
                        Paragraph(str(re.get('ErrorDescription', '')), style),
                    ])
 
            elif Status == "QRY":
                y_ax = 430
                filename = "QUERY_STATUS"
                TitleStatus = 'QUERY MESSAGE'
                rows = SqlDb.execute_query(
                    """SELECT * FROM CommonRejectStatus
                       WHERE MsgId = %s AND MailBoxId = %s
                       ORDER BY Sno""",
                    [MSGId, MailId]
                )
                for re in rows:
                    rejectStatus.append([
                        Paragraph(str(re.get('Sno', '')), style),
                        Paragraph(str(re.get('ErrorId', '')), style),
                        Paragraph(str(re.get('ErrorDescription', '')), style),
                    ])
 
            elif Status == "CNL":
                filename = f"Cancellation_{Inheader.get('PermitNumber', '')}"
                TitleStatus = 'CANCELLATION MESSAGE'
                can.rect(50, 460, 500, 50)
 
                cancels = SqlDb.execute_query(
                    "SELECT * FROM CommonCancel WHERE Permitno = %s",
                    [Inheader.get('PermitNumber', '')]
                )
                InpayCancel = cancels[0] if cancels else {}
 
                can.drawString(
                    60, 490,
                    f"CANCELLATION REASON : {str(InpayCancel.get('ReasonForCancel', '')).upper()[:4]}"
                )
                can.drawString(
                    60, 475,
                    f"DESCRIPTION OF REASON : {str(InpayCancel.get('DescriptionOfReason', '')).upper()}"
                )
                y_ax = 370
 
                cancel_rows = SqlDb.execute_query(
                    """SELECT * FROM CommonAMDPMT
                       WHERE PermitNumber = %s
                       ORDER BY Sno""",
                    [MSGId]
                )
                for re in cancel_rows:
                    rejectStatus.append([
                        Paragraph(str(re.get('Sno', '')), style),
                        Paragraph(str(re.get('ConditionCode', '')), style),
                        Paragraph(str(re.get('ConditionDesc', '')), style),
                    ])
 
            else:
                y_ax = 430
                filename = f"Status_{PermitId}"
                TitleStatus = f'{Status} MESSAGE'
 
            # ── Title ─────────────────────────────────────────────────────
            can.setFont('Times-Bold', 10)
            can.drawString(250, 790, TitleStatus)
            can.setFont('Times-Roman', 9)
 
            # ── Transport fields ──────────────────────────────────────────
            coveyance   = ''
            Obl_hawb    = ''
            InTransportId = ''
            inwardMode = str(Inheader.get('InwardTransportMode', '') or '')[4:].upper()
 
            if inwardMode == "SEA":
                Obl_hawb      = Inheader.get('OceanBillofLadingNo', '')
                coveyance     = Inheader.get('VoyageNumber', '')
                InTransportId = Inheader.get('VesselName', '')
            elif inwardMode == "AIR":
                Obl_hawb      = Inheader.get('MasterAirwayBill', '')
                coveyance     = Inheader.get('FlightNO', '')
                InTransportId = Inheader.get('VesselName', '')
 
            # ── Info box ──────────────────────────────────────────────────
            can.rect(50, 530, 500, 250)
 
            can.drawString(60, 760, "MESSAGE TYPE : IN-PAYMENT DECLARATION ")
            can.drawString(60, 745, f"DECLARATION TYPE : {Inheader.get('DeclarationType', '')}")
            can.drawString(60, 730, f"COMPANY UEN : {Import.get('CRUEI', '')}")
            can.drawString(60, 715, f"COMPANY NAME : {Import.get('Name', '')}")
            can.drawString(60, 700, f"PORT OF LOADING : {Loading.get('portname', '')}")
            can.drawString(60, 685, "PORT OF DISCHARGE : ")
            can.drawString(60, 670, "DESTINATION COUNTRY : ")
            can.drawString(60, 655, f"IN TRANSPORT ID : {InTransportId}")
 
            # ArrivalDate
            arrDate = Inheader.get('ArrivalDate')
            try:
                arrDateStr = arrDate.strftime('%d-%m-%Y') if arrDate else ''
            except Exception:
                try:
                    arrDateStr = datetime.strptime(str(arrDate), '%Y-%m-%d').strftime('%d-%m-%Y')
                except Exception:
                    arrDateStr = str(arrDate) if arrDate else ''
 
            can.drawString(60, 640, f"ARRIVAL DATE : {arrDateStr}")
            can.drawString(60, 625, f"CONVEYANCE NO : {coveyance}")
            can.drawString(60, 610, f"MAWB / OBL : {Obl_hawb}")
            can.drawString(60, 595, "OUT TRANSPORT ID :")
            can.drawString(60, 580, "DEPARTURE DATE :")
            can.drawString(60, 565, "CONVEYANCE NO :")
            can.drawString(60, 550, "MAWB / OBL :")
 
            # TouchTime / create date
            touchTime = Inheader.get('TouchTime')
            try:
                createDate = touchTime.strftime("%d-%m-%Y") if touchTime else ''
            except Exception:
                createDate = str(touchTime)[:10] if touchTime else ''
 
            can.drawString(310, 760, f"CREATE DATE : {createDate}")
            can.drawString(310, 745, f"MESSAGE ID : {MSGId}")
            can.drawString(310, 730, f"PERMIT NUMBER : {Inheader.get('PermitNumber', '')}")
            can.drawString(310, 715, f"PREVIOUS PERMIT : {Inheader.get('PreviousPermit', '')}")
            can.drawString(310, 700, f"PLACE OF RELEASE : {Inheader.get('ResLoaName', '') or Inheader.get('ReleaseLocName', '')}")
            can.drawString(310, 685, f"PLACE OF RECEIPT : {Inheader.get('RecepitLocName', '')}")
            can.drawString(310, 670, f"TOTAL OUTER PACKAGE : {Inheader.get('TotalOuterPack', '')}/{Inheader.get('TotalOuterPackUOM', '')}")
            can.drawString(310, 655, f"TOTAL GROSS WEIGHT : {Inheader.get('TotalGrossWeight', '')}/{Inheader.get('TotalGrossWeightUOM', '')}")
 
            # Item count
            item_count_rows = SqlDb.execute_query(
                "SELECT COUNT(*) AS cnt FROM CommonItemDtl WHERE PermitId = %s", [PermitId]
            )
            item_count = item_count_rows[0]['cnt'] if item_count_rows else 0
            can.drawString(310, 640, f"NUMBER OF ITEMS : {item_count}")
 
            # ── Status table ──────────────────────────────────────────────
            col_widths = [90, 70, 340]
            table = Table(rejectStatus, colWidths=col_widths)
            table.setStyle(TableStyle([
                ('FONTNAME',       (0, 0), (-1, -1), 'Times-Roman'),
                ('FONTSIZE',       (0, 0), (-1, -1), 9),
                ('GRID',           (0, 0), (-1, -1), 1, colors.black),
                ('BOX',            (0, 0), (-1, -1), 0.25, colors.black),
                ('VALIGN',         (0, 0), (-1, -1), 'TOP'),
            ]))
            table.wrapOn(can, 0, 0)
            table.drawOn(can, 50, y_ax)
 
            can.showPage()
            can.save()
 
            packet.seek(0)
            response = HttpResponse(packet, content_type='application/pdf')
            response['Content-Disposition'] = (
                f'attachment; filename="{filename}{datetime.now().strftime("%d-%m-%Y_%H%M")}.pdf"'
            )
            return response
 
        except Exception as e:
            import traceback
            traceback.print_exc()
            return HttpResponse(f"Error: {str(e)}", status=500)

# Gst status
class GstStatus(APIView):
    table = "CommonHeaderTbl"

    def get(self, request):
        try:
            permit_id = request.GET.get("PermitId")
            if not permit_id:
                return Response(
                    {"error": "PermitId is required"},
                    status=status.HTTP_400_BAD_REQUEST
                )

            # Fetch current status
            rows = SqlDb.execute_query(
                f"SELECT Status FROM {self.table} WHERE PermitId = %s",
                [permit_id]
            )

            if not rows:
                return Response(
                    {"error": f"No record found for PermitId {permit_id}"},
                    status=status.HTTP_404_NOT_FOUND
                )

            current_status = rows[0].get("Status", "")

            # Only update if current status is WFA
            if current_status != "WFA":
                return Response(
                    {
                        "error": f"Status is '{current_status}'. Only WFA status can be changed to NEW.",
                        "Status": current_status
                    },
                    status=status.HTTP_400_BAD_REQUEST
                )

            # Update status to NEW
            SqlDb.execute_query(
                f"UPDATE {self.table} SET Status = 'NEW' WHERE PermitId = %s",
                [permit_id]
            )
            SqlDb.commit()

            return Response(
                {
                    "message": f"Status updated from WFA to NEW for PermitId {permit_id}",
                    "PermitId": permit_id,
                    "Status": "NEW"
                },
                status=status.HTTP_200_OK
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response(
                {"error": str(e)},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )


class GetPermitConditions(APIView):
    def get(self, request):
        try:
            permit_id = request.GET.get("PermitId")
            if not permit_id:
                return Response({"error": "PermitId is required"}, status=400)

            # Get header info
            header_rows = SqlDb.execute_query(
                "SELECT MSGId, TradeNetMailboxID, Status, JobId, PermitNumber, TouchTime, TouchUser FROM CommonHeaderTbl WHERE PermitId = %s",
                [permit_id]
            )
            if not header_rows:
                return Response({"error": "Permit not found"}, status=404)

            header = header_rows[0]
            msg_id = header.get("MSGId", "")
            mail_id = header.get("TradeNetMailboxID", "")
            status_val = header.get("Status", "")

            records = []

            if status_val == "ERR":
                rows = SqlDb.execute_query(
                    "SELECT * FROM CommonErrorStatus WHERE MsgId = %s ORDER BY Sno",
                    [msg_id]
                )
                for row in rows:
                    records.append({
                        "Sno": row.get("Sno"),
                        "ErrorCode": row.get("ErrorId") or row.get("ErrorCode") or "",
                        "Description": row.get("ErrorDescription") or row.get("Description") or "",
                    })
            else:
                rows = SqlDb.execute_query(
                    "SELECT * FROM CommonRejectStatus WHERE MsgId = %s AND MailBoxId = %s ORDER BY Sno",
                    [msg_id, mail_id]
                )
                for row in rows:
                    records.append({
                        "Sno": row.get("Sno"),
                        "Code": row.get("ErrorId") or row.get("Code") or "",
                        "Description": row.get("ErrorDescription") or row.get("Description") or "",
                    })

            touch_time = header.get("TouchTime")
            dec_date = touch_time.strftime("%d/%m/%Y %H:%M:%S") if touch_time else ""

            return Response({
                "PermitNo": header.get("PermitNumber") or "-",
                "ValidityPeriod": "-",
                "PermitApprovedDate": "-",
                "JobId": header.get("JobId") or "-",
                "MsgId": msg_id,
                "DecDate": dec_date,
                "SubmittedBy": header.get("TouchUser") or "-",
                "Status": status_val,
                "TransmitUser": "",
                "CreatedBy": header.get("TouchUser") or "-",
                "Records": records,
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)


# class TransmitInnonpayment(APIView):
#     def post(self, request):
#         try:
#             permit_ids       = request.data.get("permitIds", [])
#             declaration_type = request.data.get("declarationType", "")
#             username         = request.data.get("user") or request.session.get("Username")
#             touch_time       = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

#             # ── Validation ─────────────────────────────────────────────────
#             if not permit_ids:
#                 return Response({"error": "No permits selected"}, status=400)
#             if not declaration_type:
#                 return Response({"error": "Declaration type is required"}, status=400)
#             if not username:
#                 return Response({"error": "User is required"}, status=400)

#             now        = datetime.now()
#             ref_date   = now.strftime("%Y%m%d")
#             job_date   = now.strftime("%y%m%d")
#             today_dash = now.strftime("%Y-%m-%d")

#             copied_permits = []

#             with transaction.atomic():
#                 cursor = connection.cursor()

#                 # ── Get user account details ───────────────────────────────
#                 cursor.execute(
#                     "SELECT AccountId, MailBoxId FROM ManageUser WHERE UserName = %s",
#                     [username]
#                 )
#                 row = cursor.fetchone()
#                 if not row:
#                     return Response({"error": f"User '{username}' not found"}, status=404)
#                 account_id, mailbox_id = row

#                 for permit_id in permit_ids:

#                     # ── Check permit exists ────────────────────────────────
#                     cursor.execute(
#                         "SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s",
#                         [permit_id]
#                     )
#                     if not cursor.fetchone():
#                         continue

#                     # ── Generate new IDs ───────────────────────────────────
#                     cursor.execute(
#                         """
#                         SELECT ISNULL(COUNT(*), 0) + 1
#                         FROM CommonHeaderTbl
#                         WHERE MSGId LIKE %s AND MessageType = 'INPDEC'
#                         """,
#                         [f"%{ref_date}%"]
#                     )
#                     ref_count = cursor.fetchone()[0]
#                     ref_id = f"{ref_count:03d}"

#                     cursor.execute(
#                         """
#                         SELECT ISNULL(COUNT(*), 0) + 1
#                         FROM PermitCount
#                         WHERE TouchTime LIKE %s AND AccountId = %s
#                         """,
#                         [f"%{today_dash}%", account_id]
#                     )
#                     job_count     = cursor.fetchone()[0]
#                     job_id        = f"K{job_date}{job_count:05d}"
#                     msg_id        = f"{ref_date}{job_count:04d}"
#                     new_permit_id = f"{username}{ref_date}{ref_id}"

#                     cursor.execute(
#                         """
#                         INSERT INTO CommonHeaderTbl (
#                             Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
#                             MessageType, DeclarationType,
#                             PreviousPermit, CargoPackType,
#                             InwardTransportMode, OutwardTransportMode, BGIndicator,
#                             SupplyIndicator, ReferenceDocuments, License,
#                             COType, Entryyear, GSPDonorCountry,
#                             CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
#                             PerCommon, CurrencyCode, AddCerDtl, TransDtl,
#                             Recipient, DeclarantCompanyCode, ExporterCompanyCode,
#                             Inwardcarriercode, OutwardCarrierAgentCode,
#                             FreightForwarderCode, ImporterCompanyCode,
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
#                             'INPDEC',%s,
#                             PreviousPermit, CargoPackType,
#                             InwardTransportMode, OutwardTransportMode, BGIndicator,
#                             SupplyIndicator, ReferenceDocuments, License,
#                             COType, Entryyear, GSPDonorCountry,
#                             CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
#                             PerCommon, CurrencyCode, AddCerDtl, TransDtl,
#                             Recipient, DeclarantCompanyCode, ExporterCompanyCode,
#                             NULL, NULL,
#                             FreightForwarderCode, ImporterCompanyCode,
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
#                             Defrentprinting, Cnb,
#                             NULL,           -- DeclarningFor: NULL to avoid truncation error
#                             MRDate, MRTime,
#                             CondColor, TransmitId, gstVerified
#                         FROM CommonHeaderTbl
#                         WHERE PermitId = %s
#                         """,
#                         [
#                             ref_id, job_id, msg_id, new_permit_id, mailbox_id,
#                             declaration_type,
#                             username, touch_time,
#                             permit_id 
#                         ]
#                     )

#                     # ── Copy all child tables (same as CopyInpayment) ──────
#                     child_tables = {
#                         "CommonInvoiceDtl": [
#                             "SNo", "InvoiceNo", "InvoiceDate", "TermType",
#                             "AdValoremIndicator", "PreDutyRateIndicator",
#                             "SupplierImporterRelationship", "SupplierCode", "ImportPartyCode",
#                             "TICurrency", "TIExRate", "TIAmount", "TISAmount",
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
#                             "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
#                             "ImPQty", "ImPUOM", "PreferentialCode",
#                             "GSTRate", "GSTUOM", "GSTAmount",
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
#                             "ItemNo", "ProductCode", "Quantity", "ProductUOM",
#                             "RowNo", "CascCode1", "CascCode2", "CascCode3",
#                             "MessageType", "TouchUser", "TouchTime", "EndUserDes", "CASCId"
#                         ],
#                         "CommonCPCDtl": [
#                             "MessageType", "RowNo", "CPCType",
#                             "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
#                             "TouchUser", "TouchTime"
#                         ],
#                         "CommonContainerDtl": [
#                             "RowNo", "ContainerNo", "Size", "Weight",
#                             "SealNo", "MessageType", "TouchUser", "TouchTime"
#                         ],
#                         "CommonFile": [
#                             "Name", "ContentType", "Data", "DocumentType",
#                             "TouchUser", "TouchTime", "filePath", "Size", "Type"
#                         ],
#                     }

#                     for table, cols in child_tables.items():
#                         col_list    = ", ".join(["PermitId"] + cols)
#                         select_cols = ", ".join(cols)
#                         try:
#                             cursor.execute(
#                                 f"""
#                                 INSERT INTO {table} ({col_list})
#                                 SELECT %s, {select_cols}
#                                 FROM {table}
#                                 WHERE PermitId = %s
#                                 """,
#                                 [new_permit_id, permit_id]
#                             )
#                             if "MessageType" in cols:
#                                 cursor.execute(
#                                     f"UPDATE {table} SET MessageType = 'INPDEC' WHERE PermitId = %s",
#                                     [new_permit_id]
#                                 )

#                         except Exception as err:
#                             print(f"Warning copying {table}: {err}")

#                     # ── Insert PermitCount ─────────────────────────────────
#                     cursor.execute(
#                         """
#                         INSERT INTO PermitCount
#                             (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
#                         VALUES (%s, 'INPDEC', %s, %s, %s, %s)
#                         """,
#                         [new_permit_id, account_id, msg_id, username, touch_time]
#                     )

#                     copied_permits.append(new_permit_id)

#             return Response({
#                 "SUCCESS": True,
#                 "message": f"{len(copied_permits)} permit(s) transmitted successfully as InNon Payment",
#                 "copiedPermits": copied_permits,
#                 "declarationType": declaration_type,
#                 "messageType": "INPDEC",
#             })

#         except Exception as e:
#             import traceback
#             traceback.print_exc()
#             return Response({"error": str(e)}, status=500)

class TransmitInnonpayment(APIView):
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

            now        = datetime.now()
            ref_date   = now.strftime("%Y%m%d")
            job_date   = now.strftime("%y%m%d")
            today_dash = now.strftime("%Y-%m-%d")

            copied_permits = []

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

                # GLOBAL starting count for JobId/MsgId (matches CopyInpayment/PostCommonHeaderTable)
                cursor.execute(
                    """
                    SELECT ISNULL(COUNT(*), 0) + 1 AS Count
                    FROM CommonHeaderTbl
                    WHERE JobId LIKE %s
                    """,
                    [f"K{job_date}%"]
                )
                job_count = cursor.fetchone()[0]

                # GLOBAL starting count for PermitId/RefId (per-user, today)
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
                        continue

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
                            FreightForwarderCode, ImporterCompanyCode,
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
                            'INPDEC',%s,
                            PreviousPermit, CargoPackType,
                            InwardTransportMode, OutwardTransportMode, BGIndicator,
                            SupplyIndicator, ReferenceDocuments, License,
                            COType, Entryyear, GSPDonorCountry,
                            CerDetailtype1, CerDetailCopies1, CerDetailtype2, CerDetailCopies2,
                            PerCommon, CurrencyCode, AddCerDtl, TransDtl,
                            Recipient, DeclarantCompanyCode, ExporterCompanyCode,
                            NULL, NULL,
                            FreightForwarderCode, ImporterCompanyCode,
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

                    # ── Copy all child tables (same as CopyInpayment) ──────
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
                            if "MessageType" in cols:
                                cursor.execute(
                                    f"UPDATE {table} SET MessageType = 'INPDEC' WHERE PermitId = %s",
                                    [new_permit_id]
                                )

                        except Exception as err:
                            print(f"Warning copying {table}: {err}")

                    # ── Insert PermitCount ─────────────────────────────────
                    cursor.execute(
                        """
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'INPDEC', %s, %s, %s, %s)
                        """,
                        [new_permit_id, account_id, msg_id, username, touch_time]
                    )

                    copied_permits.append(new_permit_id)

                    # Increment counters for next permit in this batch
                    job_count += 1
                    ref_count += 1

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) transmitted successfully as InNon Payment",
                "copiedPermits": copied_permits,
                "declarationType": declaration_type,
                "messageType": "INPDEC",
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)

# MAIL BOX TRANSMIT

class MailBoxTransmitData(APIView):
    def post(self, request):
        try:
            permit_ids  = request.data.get("permitIds", [])
            mailbox_id  = request.data.get("mailboxId", "")
            username    = request.data.get("user")
            touch_time  = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if not permit_ids:
                return Response({"error": "No permits selected"}, status=400)
            if not mailbox_id:
                return Response({"error": "Mailbox ID is required"}, status=400)
            if not username:
                return Response({"error": "User is required"}, status=400)

            now        = datetime.now()
            ref_date   = now.strftime("%Y%m%d")
            job_date   = now.strftime("%y%m%d")
            today_dash = now.strftime("%Y-%m-%d")

            copied_permits = []

            with transaction.atomic():
                cursor = connection.cursor()

                # Get the account for the TARGET mailbox user
                cursor.execute(
                    "SELECT AccountId FROM ManageUser WHERE MailBoxId = %s AND Status='Active'",
                    [mailbox_id]
                )
                row = cursor.fetchone()
                if not row:
                    return Response({"error": f"No active user found for mailbox '{mailbox_id}'"}, status=404)
                target_account_id = row[0]

                # Get current user's account
                cursor.execute(
                    "SELECT AccountId FROM ManageUser WHERE UserName = %s",
                    [username]
                )
                row = cursor.fetchone()
                if not row:
                    return Response({"error": f"User '{username}' not found"}, status=404)

                for permit_id in permit_ids:
                    cursor.execute(
                        "SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s",
                        [permit_id]
                    )
                    if not cursor.fetchone():
                        continue

                    # RefId
                    cursor.execute(
                        """
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM CommonHeaderTbl
                        WHERE MSGId LIKE %s AND MessageType = 'IPTDEC'
                        """,
                        [f"%{ref_date}%"]
                    )
                    ref_count  = cursor.fetchone()[0]
                    ref_id     = f"{ref_count:03d}"

                    # JobId + MsgId scoped to target account
                    cursor.execute(
                        """
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM PermitCount
                        WHERE TouchTime LIKE %s AND AccountId = %s
                        """,
                        [f"%{today_dash}%", target_account_id]
                    )
                    job_count     = cursor.fetchone()[0]
                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    # Copy header with new mailbox
                    cursor.execute(
                        """
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
                            FreightForwarderCode, ImporterCompanyCode,
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
                            FreightForwarderCode, ImporterCompanyCode,
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
                        """,
                        [
                            ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                            username, touch_time,
                            permit_id
                        ]
                    )

                    # Copy all child tables (identical to CopyInpayment)
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
                        except Exception as err:
                            print(f"Warning copying {table}: {err}")

                    # Insert PermitCount for target account
                    cursor.execute(
                        """
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
                        """,
                        [new_permit_id, target_account_id, msg_id, username, touch_time]
                    )

                    copied_permits.append(new_permit_id)

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) transmitted to mailbox '{mailbox_id}' successfully",
                "copiedPermits": copied_permits,
                "targetMailbox": mailbox_id,
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": str(e)}, status=500)


class GetExchangeRateByDate(APIView):
    table = "CustomsCurrency"

    def get(self, request):
        try:
            date = request.GET.get("date")

            if not date:
                return Response(
                    {"error": "date is required"},
                    status=status.HTTP_400_BAD_REQUEST
                )

            query = f"""
                SELECT CurrencyCode, CurrencyDescription, CurrencyExchangeRate, RateUnit, CountryName
                FROM {self.table}
                WHERE CONVERT(date, %s, 23) BETWEEN StartDate AND EndDate
                ORDER BY CurrencyCode
            """
            rows = SqlDb.execute_query(query, [date])

            if not rows:
                return Response(
                    {"error": "No exchange rate data found for this date"},
                    status=status.HTTP_404_NOT_FOUND
                )

            records = [
                {
                    "Currency": row["CurrencyCode"],
                    "Description": row["CurrencyDescription"],
                    "Rate": row["CurrencyExchangeRate"],
                    "RateUnit": row["RateUnit"],
                    "Country": row["CountryName"],
                }
                for row in rows
            ]

            return Response(
                {
                    "Date": date,
                    "Records": records,
                },
                status=status.HTTP_200_OK
            )

        except Exception as e:
            return Response(
                {"error": str(e)},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )

            
class CopyRefund(APIView):

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

                    # ── 1. Insert new CommonHeaderTbl row — copy everything as-is ──
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
                            CondColor, TransmitId, gstVerified, Message,
                            HandlingAgentCode, CustomerRemarks
                        )
                        SELECT
                            %s, %s, %s, %s, TradeNetMailboxID,
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
                            'DRF', %s, %s,
                            PermitNumber, 'RFD',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified, Message,
                            HandlingAgentCode, CustomerRemarks
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror the row we just inserted into InHeaderTbl ──
                    # Reads from the NEW CommonHeaderTbl row (new_permit_id), so no
                    # recomputation of Ref/Job/Msg ids or dates — single source of truth.
                    cursor.execute("""
                        INSERT INTO InHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ReleaseLocName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
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
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: copy into BOTH Common* and In-payment* targets ──
                    child_tables = {
                        "CommonInvoiceDtl": ("InvoiceDtl", [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
                        ]),
                        "CommonItemDtl": ("ItemDtl", [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "Brand", "Model", "InHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode", "GSTRate", "GSTUOM",
                            "GSTAmount", "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "LSPValue", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                        ]),
                        "CommonCASCDtl": ("CASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM", "RowNo",
                            "CascCode1", "CascCode2", "CascCode3", "MessageType",
                            "TouchUser", "TouchTime", "CASCId"
                        ]),
                        "CommonCPCDtl": ("CPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("ContainerDtl", [
                            "RowNo", "ContainerNo", "Size", "Weight", "SealNo",
                            "MessageType", "TouchUser", "TouchTime"
                        ]),
                    }

                    for src_table, (dst_table, cols) in child_tables.items():
                        col_list = ", ".join(cols)
                        # Common -> Common (from OLD permit, full column set)
                        try:
                            cursor.execute(
                                f"""INSERT INTO {src_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, permit_id]
                            )
                        except Exception as err:
                            print(f"Warning copying {src_table}: {err}")

                        # Common(NEW) -> In-payment target — single read of what we just
                        # wrote, so both tables end up byte-identical for this permit.
                        try:
                            cursor.execute(
                                f"""INSERT INTO {dst_table} (PermitId, {col_list})
                                    SELECT %s, {col_list} FROM {src_table} WHERE PermitId = %s""",
                                [new_permit_id, new_permit_id]
                            )
                        except Exception as err:
                            print(f"Warning mirroring {src_table} -> {dst_table}: {err}")

                    # ── 4. File & PMT — column names differ, handled explicitly ──
                    try:
                        cursor.execute("""
                            INSERT INTO CommonFile (PermitId, Sno,Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                            SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
                            FROM CommonFile WHERE PermitId = %s
                        """, [new_permit_id, permit_id])
                        cursor.execute("""
                            INSERT INTO InFile (PermitId,Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                                SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
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
                        # InPMT has a different schema (Sno, CertificateNumber, StartDate, etc.)
                        # — populated on approval, not at copy time, same as CopyInpayment.
                    except Exception as err:
                        print(f"Warning copying PMT: {err}")

                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
                    """, [new_permit_id, account_id, msg_id, username, touch_time])

                    copied_permits.append(new_permit_id)
                    job_count += 1
                    ref_count += 1

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) copied for refund successfully",
                "copiedPermits": copied_permits,
            })
        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({"error": f"Database Error: {str(e)}"}, status=400)

# copy amend

class CopyAmend(APIView):

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

                    # ── 1. Insert new CommonHeaderTbl row — copy everything as-is ──
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
                            CondColor, TransmitId, gstVerified, Message,
                            HandlingAgentCode, CustomerRemarks
                        )
                        SELECT
                            %s, %s, %s, %s, TradeNetMailboxID,
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
                            'DRF', %s, %s,
                            PermitNumber, 'AMD',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified, Message,
                            HandlingAgentCode, CustomerRemarks
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror into InHeaderTbl ──
                    cursor.execute("""
                        INSERT INTO InHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ReleaseLocName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
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
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: copy into BOTH Common* and In-payment* targets ──
                    child_tables = {
                        "CommonInvoiceDtl": ("InvoiceDtl", [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
                        ]),
                        "CommonItemDtl": ("ItemDtl", [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "Brand", "Model", "InHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode", "GSTRate", "GSTUOM",
                            "GSTAmount", "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "LSPValue", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                        ]),
                        "CommonCASCDtl": ("CASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM", "RowNo",
                            "CascCode1", "CascCode2", "CascCode3", "MessageType",
                            "TouchUser", "TouchTime", "CASCId"
                        ]),
                        "CommonCPCDtl": ("CPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("ContainerDtl", [
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
                            INSERT INTO InFile (PermitId,Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                                SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
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
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
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


# copy cancel
class CopyCancel(APIView):

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

                    # ── 1. Insert new CommonHeaderTbl row — copy everything as-is ──
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
                            CondColor, TransmitId, gstVerified, Message,
                            HandlingAgentCode, CustomerRemarks
                        )
                        SELECT
                            %s, %s, %s, %s, TradeNetMailboxID,
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
                            'DRF', %s, %s,
                            PermitNumber, 'CNL',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified, Message,
                            HandlingAgentCode, CustomerRemarks
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id,
                        username, touch_time,
                        permit_id
                    ])

                    # ── 2. Mirror into InHeaderTbl ──
                    cursor.execute("""
                        INSERT INTO InHeaderTbl (
                            Refid, JobId, MSGId, PermitId, TradeNetMailboxID,
                            MessageType, DeclarationType, PreviousPermit, CargoPackType,
                            InwardTransportMode, BGIndicator, SupplyIndicator,
                            ReferenceDocuments, License, Recipient, DeclarantCompanyCode,
                            ImporterCompanyCode, InwardCarrierAgentCode, FreightForwarderCode,
                            ClaimantPartyCode, HBL, ArrivalDate, LoadingPortCode,
                            VoyageNumber, VesselName, OceanBillofLadingNo, ConveyanceRefNo,
                            TransportId, FlightNO, AircraftRegNo, MasterAirwayBill,
                            ReleaseLocation, ReleaseLocName, RecepitLocation,
                            TotalOuterPack, TotalOuterPackUOM,
                            TotalGrossWeight, TotalGrossWeightUOM,
                            GrossReference, BlanketStartDate, TradeRemarks, InternalRemarks,
                            DeclareIndicator, NumberOfItems,
                            TotalCIFFOBValue, TotalGSTTaxAmt, TotalExDutyAmt,
                            TotalCusDutyAmt, TotalODutyAmt, TotalAmtPay,
                            Status, TouchUser, TouchTime, PermitNumber, prmtStatus,
                            RecepitLocName, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
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
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [new_permit_id])

                    # ── 3. Child tables: copy into BOTH Common* and In-payment* targets ──
                    child_tables = {
                        "CommonInvoiceDtl": ("InvoiceDtl", [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
                            "OTCCharge", "OTCCurrency", "OTCExRate", "OTCAmount", "OTCSAmount",
                            "FCCharge", "FCCurrency", "FCExRate", "FCAmount", "FCSAmount",
                            "ICCharge", "ICCurrency", "ICExRate", "ICAmount", "ICSAmount",
                            "CIFSUMAmount", "GSTPercentage", "GSTSUMAmount",
                            "MessageType", "TouchUser", "TouchTime", "ChkOtherInv"
                        ]),
                        "CommonItemDtl": ("ItemDtl", [
                            "ItemNo", "MessageType", "HSCode", "Description", "DGIndicator",
                            "Contry", "Brand", "Model", "InHAWBOBL", "DutiableQty", "DutiableUOM",
                            "TotalDutiableQty", "TotalDutiableUOM", "InvoiceQuantity",
                            "HSQty", "HSUOM", "AlcoholPer", "InvoiceNo",
                            "ChkUnitPrice", "UnitPrice", "UnitPriceCurrency", "ExchangeRate",
                            "SumExchangeRate", "TotalLineAmount", "InvoiceCharges", "CIFFOB",
                            "OPQty", "OPUOM", "IPQty", "IPUOM", "InPqty", "InPUOM",
                            "ImPQty", "ImPUOM", "PreferentialCode", "GSTRate", "GSTUOM",
                            "GSTAmount", "ExciseDutyRate", "ExciseDutyUOM", "ExciseDutyAmount",
                            "CustomsDutyRate", "CustomsDutyUOM", "CustomsDutyAmount",
                            "OtherTaxRate", "OtherTaxUOM", "OtherTaxAmount",
                            "CurrentLot", "PreviousLot", "LSPValue", "Making",
                            "ShippingMarks1", "ShippingMarks2", "ShippingMarks3", "ShippingMarks4",
                            "TouchUser", "TouchTime", "VehicleType", "OptionalChrgeUOM",
                            "EngineCapcity", "Optioncahrge", "OptionalSumtotal",
                            "OptionalSumExchage", "EngineCapUOM", "orignaldatereg"
                        ]),
                        "CommonCASCDtl": ("CASCDtl", [
                            "ItemNo", "ProductCode", "Quantity", "ProductUOM", "RowNo",
                            "CascCode1", "CascCode2", "CascCode3", "MessageType",
                            "TouchUser", "TouchTime", "CASCId"
                        ]),
                        "CommonCPCDtl": ("CPCDtl", [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
                        ]),
                        "CommonContainerDtl": ("ContainerDtl", [
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
                            INSERT INTO InFile (PermitId,Sno, Name, ContentType, Data,
                                DocumentType, TouchUser, TouchTime, filePath, Size, Type)
                                SELECT %s, Sno,Name, ContentType, Data, DocumentType,
                                TouchUser, TouchTime, filePath, Size, Type
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
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
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