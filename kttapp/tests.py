from django.db import connections
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from django.http import HttpResponse
from django.http import FileResponse
from django.conf import settings
import os


"""The below sqldb is used for commanly we can use everywhere if we need
 database's table just create new class and use inherit"""
class SqlDb:
    database_name = "default"
    @classmethod
    def execute_query(cls, query, params=None):
        """---Class method to execute SELECT queries and return list of dicts---"""
        with connections[cls.database_name].cursor() as cursor:
            cursor.execute(query, params or [])
            columns = [col[0] for col in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall()]

#18-02-26
#--------Login Page-----

class LoginAPIView(APIView):
    def post(self,request):
        username = request.data.get("Username")
        password = request.data.get("Password")
        check = request.data.get("ChkLogin")
        if not username or not password:
            return Response({"error": "Username and Password are required"}, status=status.HTTP_400_BAD_REQUEST)
        query = "SELECT * FROM ManageUser WHERE UserName = %s"
        users = SqlDb.execute_query(query, [username])
        if not users:
            return Response({"error": "Enter Correct Password and Username"}, status=status.HTTP_401_UNAUTHORIZED)
        user = users[0]
        if user["Password"] != password:
            return Response({"error": "Enter Correct Password and Username"}, status=status.HTTP_401_UNAUTHORIZED)
        if user["LoginStatus"] == "True":
            if check:
                request.session["Username"] = username
                request.session["Permit_Id"] = "NEW"
                return Response({"success": "Success"})
            else:
                return Response({"error": "User Already Login with another System"}, status=status.HTTP_403_FORBIDDEN)
        update_query = "UPDATE ManageUser SET LoginStatus='True' WHERE UserName=%s"
        SqlDb.execute_query(update_query, [username])
        request.session["Username"] = username
        request.session["Permit_Id"] = "NEW"
        return Response({"success": "Success"})

"""The below code has used for getting all datas from table using restapi format"""
#05-02-26
#-------Header Table-------
class GetCommonHeaderTable(APIView):
    table = "CommonHeaderTbl"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT * FROM {self.table}")
        return Response(data)

class GetCommonHeaderByPermit(APIView):
    table = "CommonHeaderTbl"
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

class PostCommonHeaderTable(APIView):
    table = "CommonHeaderTbl"

    def post(self, request):
        payloads = request.data
        if not payloads:
            return Response({"error": "No data provided"}, status=400)

        if not isinstance(payloads, list):
            payloads = [payloads]

        inserted_count = 0
        for item in payloads:
            if not isinstance(item, dict):
                item = dict(item) 
            item.pop("Id", None) 
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

# class GetInvoiceByPermitId(APIView):
#     table = "CommonInvoiceDtl"

#     def get(self, request, permit_id):

#         query = f"SELECT InvoiceNo,TICurrency,TIExRate FROM {self.table} WHERE PermitId = %s"
#         data = SqlDb.execute_query(query, [permit_id])

#         if not data:
#             return Response(
#                 {"message": f"No invoices found for PermitId {permit_id}"},
#                 status=status.HTTP_404_NOT_FOUND
#             )

#         return Response(data)

class GetInvoiceByPermitId(APIView):
    table = "CommonInvoiceDtl"

    def get(self, request, permit_id):

        query = f"SELECT * FROM {self.table} WHERE PermitId = %s"
        data = SqlDb.execute_query(query, [permit_id])

        if not data:
            return Response(
                {"message": f"No invoices found for PermitId {permit_id}"},
                status=status.HTTP_404_NOT_FOUND
            )

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

class DeleteItem(APIView):
    table = "CommonItemDtl"
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
            "Result": "Item deleted successfully",
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
                    if not item_no or not permit_id or row_no is None:
                        continue
                    cursor.execute("""
                        SELECT COUNT(*)
                        FROM CommonCASCDtl
                        WHERE ItemNo=%s AND PermitId=%s AND RowNo=%s
                    """, [item_no, permit_id, row_no])
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
                                EndUserDes=%s
                            WHERE ItemNo=%s AND PermitId=%s AND RowNo=%s
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
                            item_no,
                            permit_id,
                            row_no
                        ])
                        updated_count += 1
                    else:
                        cursor.execute("""
                            INSERT INTO CommonCASCDtl
                            (
                                ItemNo, ProductCode, Quantity, ProductUOM,
                                RowNo, CascCode1, CascCode2, CascCode3,
                                PermitId, MessageType,
                                TouchUser, TouchTime, EndUserDes
                            )
                            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
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
                            item.get("EndUserDes")
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

class PostCpcTable(APIView):
    table ="CommonCPCDtl"
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


class PostFileTable(APIView):
    table ="CommonFile"
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

#DeclarationType--inpaymetn
class GetDeclarationTypeFromCommonMasterForInpayment(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId= 1 AND StatusID=1 ORDER BY Name")
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

#VesselType
class GetVesselTypeFromCommonMaster(APIView):
    table = "CommonMaster"
    def get(self, request):
        data = SqlDb.execute_query(f"SELECT Name from {self.table} WHERE TypeId=17 AND StatusID=1 ORDER BY Name")
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
        ItemInfo = pd.read_excel(xlsx_file, sheet_name="ItemInfo").fillna('')
        CascInfo = pd.read_excel(xlsx_file, sheet_name="Casccodes").fillna('')

        # =========================================================
        # ITEM INSERT CONFIG
        # =========================================================
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
            "VehicleType",
            "OptionalChrgeUOM",
            "EngineCapcity", "EngineCapUOM",
            "Optioncahrge", "OptionalSumtotal", "OptionalSumExchage",
            "orignaldatereg"
        ]

        ITEM_SQL = f"""
            INSERT INTO CommonItemDtl ({",".join(ITEM_COLUMNS)})
            VALUES ({",".join(["%s"] * len(ITEM_COLUMNS))})
        """

        # =========================================================
        # INSERT ITEMS
        # =========================================================
        item_no = 0

        for _, row in ItemInfo.iterrows():
            item_no += 1

            values = [
                item_no,
                permit_id,
                msg_type,

                row.get("HSCode"),
                row.get("Description"),
                row.get("DGIndicator"),
                row.get("Contry"),

                row.get("EndUserDescription"),
                row.get("Brand"),
                row.get("Model"),

                row.get("InHAWBOBL"),
                row.get("OutHAWBOBL"),

                row.get("DutiableQty", "0.00"),
                row.get("DutiableUOM"),

                row.get("TotalDutiableQty", "0.00"),
                row.get("TotalDutiableUOM"),

                row.get("InvoiceQuantity", "0.00"),
                row.get("HSQty", "0.00"),
                row.get("HSUOM"),

                row.get("AlcoholPer", "0.00"),
                row.get("InvoiceNo"),

                row.get("ChkUnitPrice", "0.00"),
                row.get("UnitPrice", "0.00"),
                row.get("UnitPriceCurrency"),

                row.get("ExchangeRate", "0.00"),
                row.get("SumExchangeRate", "0.00"),

                row.get("TotalLineAmount", "0.00"),
                row.get("InvoiceCharges", "0.00"),
                row.get("CIFFOB", "0.00"),

                row.get("OPQty", "0.00"),
                row.get("OPUOM"),

                row.get("IPQty", "0.00"),
                row.get("IPUOM"),

                row.get("InPqty", "0.00"),
                row.get("InPUOM"),

                row.get("ImPQty", "0.00"),
                row.get("ImPUOM"),

                row.get("PreferentialCode"),

                row.get("GSTRate", "0.00"),
                row.get("GSTUOM"),
                row.get("GSTAmount", "0.00"),

                row.get("ExciseDutyRate", "0.00"),
                row.get("ExciseDutyUOM"),
                row.get("ExciseDutyAmount", "0.00"),

                row.get("CustomsDutyRate", "0.00"),
                row.get("CustomsDutyUOM"),
                row.get("CustomsDutyAmount", "0.00"),

                row.get("OtherTaxRate", "0.00"),
                row.get("OtherTaxUOM"),
                row.get("OtherTaxAmount", "0.00"),

                row.get("CurrentLot"),
                row.get("PreviousLot"),

                row.get("LSPValue", "0.00"),
                row.get("Making"),

                row.get("ShippingMarks1"),
                row.get("ShippingMarks2"),
                row.get("ShippingMarks3"),
                row.get("ShippingMarks4"),

                row.get("CerItemQty", "0.00"),
                row.get("CerItemUOM"),
                row.get("CIFValOfCer", "0.00"),

                row.get("ManufactureCostDate"),
                row.get("TexCat"),

                row.get("TexQuotaQty", "0.00"),
                row.get("TexQuotaUOM"),

                row.get("CerInvNo"),
                row.get("CerInvDate"),

                row.get("OriginOfCer"),
                row.get("HSCodeCer"),

                row.get("PerContent"),
                row.get("CertificateDescription"),

                user_name,
                touch_time,

                row.get("VehicleType"),
                row.get("OptionalChrgeUOM"),

                row.get("EngineCapcity", "0.00"),
                row.get("EngineCapUOM"),

                row.get("Optioncahrge", "0.00"),
                row.get("OptionalSumtotal", "0.00"),
                row.get("OptionalSumExchage", "0.00"),

                row.get("orignaldatereg")
            ]

            SqlDb.execute_query(ITEM_SQL, values)

        # =========================================================
        # CASC INSERT CONFIG
        # =========================================================
        CASC_COLUMNS = [
            "ItemNo", "ProductCode", "Quantity", "ProductUOM",
            "RowNo", "CascCode1", "CascCode2", "CascCode3",
            "PermitId", "MessageType", "TouchUser", "TouchTime",
            "CASCId", "EndUserDes"
        ]

        CASC_SQL = f"""
            INSERT INTO CommonCASCDtl ({",".join(CASC_COLUMNS)})
            VALUES ({",".join(["%s"] * len(CASC_COLUMNS))})
        """

        # =========================================================
        # INSERT CASC
        # =========================================================
        for _, row in CascInfo.iterrows():

            if row.get("ProductCode"):

                casc_values = [
                    row.get("ItemNo"),
                    row.get("ProductCode"),
                    row.get("Quantity", "0.00"),
                    row.get("ProductUOM"),
                    row.get("RowNo"),
                    row.get("CascCode1"),
                    row.get("CascCode2"),
                    row.get("CascCode3"),

                    permit_id,
                    msg_type,
                    user_name,
                    touch_time,

                    row.get("CASCId"),
                    row.get("EndUserDes")
                ]

                SqlDb.execute_query(CASC_SQL, casc_values)

        # =========================================================
        # COMMIT
        # =========================================================
        SqlDb.commit()

        # =========================================================
        # RETURN DATA
        # =========================================================
        SqlDb.execute_query(
            "SELECT * FROM CommonItemDtl WHERE PermitId=%s ORDER BY ItemNo",
            (permit_id,)
        )
        items = self.cursor.fetchall()

        SqlDb.execute_query(
            "SELECT * FROM CommonCASCDtl WHERE PermitId=%s ORDER BY ItemNo",
            (permit_id,)
        )
        casc = self.cursor.fetchall()

        return Response({
            "Result": "UPLOAD SUCCESSFULLY",
            "item": items,
            "casc": casc
        })





class GetCommonHeaderByPermitId(APIView):
    def get(self, request):
        permit_id = request.GET.get("PermitId")
        if not permit_id:
            return Response({"error": "PermitId is required"}, status=400)

        rows = SqlDb.execute_query(
            "SELECT * FROM CommonHeaderTbl WHERE PermitId = %s", [permit_id]
        )
        if not rows:
            return Response({"error": "Record not found"}, status=404)

        return Response(rows[0], status=200)





















class CopyInpayment(APIView):
    def post(self, request):
        try:
            permits = request.data.get("permits", [])
            username = request.data.get("user")
            touch_time = request.data.get("touchTime") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if not permits:
                return Response({"error": "No permits selected"}, status=400)
            if not username:
                return Response({"error": "User required"}, status=400)

            now = datetime.now()
            ref_date   = now.strftime("%Y%m%d")   # 20260420
            job_date   = now.strftime("%y%m%d")   # 260420
            today_dash = now.strftime("%Y-%m-%d") # 2026-04-20

            with transaction.atomic():
                cursor = connection.cursor()

                # Get AccountId + MailBoxId
                cursor.execute("""
                    SELECT AccountId, MailBoxId
                    FROM ManageUser
                    WHERE UserName = %s
                """, [username])
                row = cursor.fetchone()
                if not row:
                    return Response({"error": f"User '{username}' not found"}, status=404)
                account_id, mailbox_id = row

                for permit_id in permits:
                    # Verify permit exists
                    cursor.execute("""
                        SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s
                    """, [permit_id])
                    if not cursor.fetchone():
                        continue

                    # RefId — count IPTDEC records for today in CommonHeaderTbl
                    cursor.execute("""
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM CommonHeaderTbl
                        WHERE MSGId LIKE %s AND MessageType = 'IPTDEC'
                    """, [f"%{ref_date}%"])
                    ref_count = cursor.fetchone()[0]
                    ref_id = f"{ref_count:03d}"

                    # JobId + MsgId — scoped to AccountId + today in PermitCount
                    cursor.execute("""
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM PermitCount
                        WHERE TouchTime LIKE %s AND AccountId = %s
                    """, [f"%{today_dash}%", account_id])
                    job_count = cursor.fetchone()[0]

                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    # Copy Header
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
                            'DRF', %s, %s,
                            NULL, 'NEW',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                        username, touch_time,
                        permit_id
                    ])

                    # Copy child tables
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
                        "CommonCPCDtl": [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
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
                            cursor.execute(f"""
                                INSERT INTO {table} ({col_list})
                                SELECT %s, {select_cols}
                                FROM {table}
                                WHERE PermitId = %s
                            """, [new_permit_id, permit_id])
                        except Exception as err:
                            print(f"Warning copying {table}: {err}")

                    # PermitCount insert
                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
                    """, [new_permit_id, account_id, msg_id, username, touch_time])

            return Response({
                "SUCCESS": True,
                "message": f"{len(permits)} permit(s) copied successfully"
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({
                "SUCCESS": False,
                "error": str(e)
            }, status=500)




class CopyInpayment(APIView):
    def post(self, request):
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

                # Get AccountId + MailBoxId
                cursor.execute("""
                    SELECT AccountId, MailBoxId
                    FROM ManageUser
                    WHERE UserName = %s
                """, [username])
                row = cursor.fetchone()  # fixed: was 'account' but then used 'row'
                if not row:
                    return Response({"error": f"User '{username}' not found"}, status=404)
                account_id, mailbox_id = row

                for permit_id in permits:
                    # Verify permit exists
                    cursor.execute("""
                        SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s
                    """, [permit_id])
                    if not cursor.fetchone():
                        continue

                    # RefId — count IPTDEC records for today
                    cursor.execute("""
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM CommonHeaderTbl
                        WHERE MSGId LIKE %s AND MessageType = 'IPTDEC'
                    """, [f"%{ref_date}%"])
                    ref_count = cursor.fetchone()[0]
                    ref_id = f"{ref_count:03d}"

                    # JobId + MsgId — scoped to AccountId + today in PermitCount
                    cursor.execute("""
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM PermitCount
                        WHERE TouchTime LIKE %s AND AccountId = %s
                    """, [f"%{today_dash}%", account_id])
                    job_count = cursor.fetchone()[0]

                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"

                    # Copy Header
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
                            'DRF', %s, %s,
                            NULL, 'NEW',
                            ResLoaName, RepLocName, RecepitLocName,
                            outHAWB, INHAWB, seastore, CertificateNumber,
                            Defrentprinting, Cnb, DeclarningFor, MRDate, MRTime,
                            CondColor, TransmitId, gstVerified
                        FROM CommonHeaderTbl
                        WHERE PermitId = %s
                    """, [
                        ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                        username, touch_time,
                        permit_id
                    ])

                    # Copy child tables
                    child_tables = {
                        "CommonInvoiceDtl": [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
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
                        ],
                        "CommonCPCDtl": [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
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
                            cursor.execute(f"""
                                INSERT INTO {table} ({col_list})
                                SELECT %s, {select_cols}
                                FROM {table}
                                WHERE PermitId = %s
                            """, [new_permit_id, permit_id])
                        except Exception as err:
                            print(f"Warning copying {table}: {err}")

                    # PermitCount insert — once per permit, after all child tables
                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
                    """, [new_permit_id, account_id, msg_id, username, touch_time])

            return Response({
                "SUCCESS": True,
                "message": f"{len(permits)} permit(s) copied successfully"
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({
                "SUCCESS": False,
                "error": str(e)
            }, status=500)




class CopyInpayment(APIView):
    def post(self, request):
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
            copied_permits = []
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
                for permit_id in permits:
                    cursor.execute("""
                        SELECT PermitId FROM CommonHeaderTbl WHERE PermitId = %s
                    """, [permit_id])
                    if not cursor.fetchone():
                        continue
                    # RefId — count IPTDEC records for today
                    cursor.execute("""
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM CommonHeaderTbl
                        WHERE MSGId LIKE %s AND MessageType = 'IPTDEC'
                    """, [f"%{ref_date}%"])
                    ref_count = cursor.fetchone()[0]
                    ref_id = f"{ref_count:03d}"
                    # JobId + MsgId — scoped to AccountId + today in PermitCount
                    cursor.execute("""
                        SELECT ISNULL(COUNT(*), 0) + 1
                        FROM PermitCount
                        WHERE TouchTime LIKE %s AND AccountId = %s
                    """, [f"%{today_dash}%", account_id])
                    job_count = cursor.fetchone()[0]
                    job_id        = f"K{job_date}{job_count:05d}"
                    msg_id        = f"{ref_date}{job_count:04d}"
                    new_permit_id = f"{username}{ref_date}{ref_id}"
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
                    """, [
                        ref_id, job_id, msg_id, new_permit_id, mailbox_id,
                        username, touch_time,
                        permit_id
                    ])
                    child_tables = {
                        "CommonInvoiceDtl": [
                            "SNo", "InvoiceNo", "InvoiceDate", "TermType",
                            "AdValoremIndicator", "PreDutyRateIndicator", "SupplierImporterRelationship",
                            "SupplierCode", "ImportPartyCode", "TICurrency", "TIExRate", "TIAmount", "TISAmount",
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
                        ],
                        "CommonCPCDtl": [
                            "MessageType", "RowNo", "CPCType",
                            "ProcessingCode1", "ProcessingCode2", "ProcessingCode3",
                            "TouchUser", "TouchTime"
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
                            cursor.execute(f"""
                                INSERT INTO {table} ({col_list})
                                SELECT %s, {select_cols}
                                FROM {table}
                                WHERE PermitId = %s
                            """, [new_permit_id, permit_id])
                        except Exception as err:
                            print(f"Warning copying {table}: {err}")

                    cursor.execute("""
                        INSERT INTO PermitCount
                            (PermitId, MessageType, AccountId, MsgId, TouchUser, TouchTime)
                        VALUES (%s, 'IPTDEC', %s, %s, %s, %s)
                    """, [new_permit_id, account_id, msg_id, username, touch_time])

                    copied_permits.append(new_permit_id)

            return Response({
                "SUCCESS": True,
                "message": f"{len(copied_permits)} permit(s) copied successfully",
                "copiedPermits": copied_permits,
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response({
                "SUCCESS": False,
                "error": str(e)
            }, status=500)



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
                    CONVERT(varchar, t1.ArrivalDate, 105) AS ETA,
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
                        WHEN t1.InwardTransportMode = '4 : Air' THEN t1.MasterAirwayBill
                        WHEN t1.InwardTransportMode = '1 : Sea' THEN t1.OceanBillofLadingNo
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
                    t1.InwardTransportMode AS TPT,
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