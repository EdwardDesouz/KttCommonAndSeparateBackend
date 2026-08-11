from django.urls import path
from . import *

urlpatterns = [
#  # Header Table
      path('getTransHeaderTableInfo/', views.GetTransHeaderTable.as_view()),
      path('getTransHeaderByPermitId/', views.GetTransHeaderByPermitId.as_view()),
      path('postTransHeaderTable/', views.PostTransHeaderTable.as_view()),
      path("editTransPermit/<str:permit_id>/", views.EditTransHeaderByPermit.as_view()),
      path("deleteTransPermit/", views.DeleteTransPermit.as_view()),

#  #CommonInvoiceTable,
#     path('outInvoiceTableInfo/', views.OutInvoiceTable.as_view()),
#     path("getOutInvoiceNo/<str:invoice_no>/<str:permit_id>/",views.GetOutInvoiceByInvoiceNo.as_view()),
#     path("getOutInvoiceByEditPermitId/",views.GetOutInvoiceByEditPermitId.as_view()),
#     path("getOutInvoiceByPermitId/<str:permit_id>/",views.GetOutInvoiceByPermitId.as_view()),
#     path("deleteOutInvoiceNo/",views.DeleteOutInnvoice.as_view()), 
#     path('postOutInvoiceTable/', views.PostOutInvoiceTable.as_view()),
#     path('editOutInvoice/<str:invoice_no>/<str:permit_id>/', views.EditOutInvoiceByInvoiceNo.as_view()),


    #CommonItemTable
    path('getTransItemTableInfo/', views.GetTransItemTabel.as_view()),
    path('getTransItemNo/<str:item_no>/<str:permit_id>/', views.GetTransItemByItemNo.as_view()),
    path("getTransItemByEditPermitId/",views.GetTransItemByEditPermitId.as_view()),
    path("deleteTransItem/",views.DeleteTransItem.as_view()),
    path('postTransItemTable/', views.PostTransItemTable.as_view()),
    path('editTransItem/<str:item_no>/<str:permit_id>/', views.EditTransItemByItemNo.as_view()),
 

# # # #   # CopyPermit
# # # #     path("copyInpayment/", views.CopyInpayment.as_view()),

    #casc
    path('getTransCascTableInfo/', views.GetTransCascTabel.as_view()),
    path('getTransCasc/<str:permit_id>/', views.GetTransCascByPermitId.as_view()),
    path('deleteTransCasc/', views.DeleteTransCasc.as_view()),
    path('deleteTransCascByCascId/<str:casc_id>/<int:row_no>/<str:permit_id>/',views.DeleteTransCascByCascId.as_view()),
    path("postTransCascTable/", views.PostTransCascTable.as_view()),
    path('editTransCasc/<str:permit_id>/', views.EditTransCascByPermitId.as_view()),
# #     #cpc
    path('getTransCpcTableInfo/', views.GetTransCpcTabel.as_view()),
    path("getTransCpcByEditPermitId/",views.GetTransCpcByEditPermitId.as_view()),
    path('getTransCpc/<str:permit_id>/', views.GetTransCpcByPermitId.as_view()),
    path("deleteTransCpc/<str:permit_id>/", views.DeleteTransCpc.as_view()),
    path("postTransCpcTable/", views.PostTransCpcTable.as_view()),
    path('editTransCpc/<str:permit_id>/', views.EditTransCpcByPermitId.as_view()),
# #     #container
    path('getTransContainerTableInfo/', views.GetTransContainerTabel.as_view()),
    path("getTransContainerByEditPermitId/",views.GetTransContainerByEditPermitId.as_view()),
    path('getTransContainer/<str:permit_id>/', views.GetTransContainerByPermitId.as_view()),
    path("deleteTransContainer/", views.DeleteTransContainer.as_view()),
    path("postTransContainerTable/", views.PostTransContainerTable.as_view()),
    path('editTransContainer/<str:permit_id>/', views.EditTransContainerByPermitId.as_view()),
# # #     #housecode
# # #     path('getInHouseItemCode/',views.GetInHouseItemCode.as_view()),
# # #     path('getInHouseItemCodeByHouseCode/<str:housecode>/',views.GetInHouseItemCodeByHouseCode.as_view()),
# # #     path('postInHouseItemCode/',views.PostInHouseItemCode.as_view()),
# # #     path('editInHouseItemCode/<str:housecode>/',views.EditInHouseItemCode.as_view()),
# # #     path('deleteInHouseItemCode/<str:housecode>/',views.DeleteInHouseItemCode.as_view()),
    #importer
    path('getTransImporterTableInfo/', views.GetTransImporterTable.as_view()),
    path('getTransImporter/<str:code>/', views.GetTransImporterByCode.as_view()),
    path('deleteTransImporter/<str:code>/', views.DeleteTransImporter.as_view()),
    path('postTransImporterTable/', views.PostTransImporterTable.as_view()),
    path('editTransImporter/<str:code>/', views.EditTransImporterByCode.as_view()),

#  HandlingAgent
   path('getTransHandlingAgentTableInfo/', views.GetTransHandlingAgentTable.as_view()),
   path('getTransHandlingAgent/<str:code>/', views.GetTransHandlingAgentByCode.as_view()),
   path("deleteTransHandlingAgent/<str:code>/", views.DeleteTransHandlingAgent.as_view()),
   path("postTransHandlingAgentTable/", views.PostTransHandlingAgentTable.as_view()),
   path('editTransHandlingAgent/<str:code>/', views.EditTransHandlingAgentByCode.as_view()),
 
    # ---------------- Freight Forwarder (FreightForwarder) ----------------
    path('getTransFreightForwarderTableInfo/', views.GetTransfreightTable.as_view()),
    path('getTransFreightForwarder/<str:code>/', views.GetTransfreightByCode.as_view()),
    path('deleteTransFreightForwarder/<str:code>/', views.DeleteTransfreight.as_view()),
    path('postTransFreightForwarderTable/', views.PostTransfreightTable.as_view()),
    path('editTransFreightForwarder/<str:code>/', views.EditTransfreightByCode.as_view()),


# #     #ClaimantParty
# #     path("getInnonClaimantPartyTable/", views.GetInnonClaimantPartyTable.as_view()),
# #     path("getInnonClaimantParty/<int:id>/", views.GetInnonClaimantPartyById.as_view()),
# #     path("deleteInnonClaimantParty/<int:id>/", views.DeleteInnonClaimantParty.as_view()),
# #     path("postInnonClaimantPartyTable/", views.PostInnonClaimantPartyTable.as_view()),
# #     path("editInnonClaimantParty/<int:id>/", views.EditInnonClaimantPartyById.as_view()), 
   
#     #Exporter
#     path('getOutExporterTableInfo/', views.GetOutExporterTable.as_view()),
#     path('getOutExporter/<str:code>/', views.GetOutExporterByCode.as_view()),
#     path("deleteOutExporter/<str:code>/", views.DeleteOutExporter.as_view()),
#     path("postOutExporterTable/", views.PostOutExporterTable.as_view()),
#     path('editOutExporter/<str:code>/', views.EditOutExporterByCode.as_view()),

# ---------------- Inward Carrier Agent (InwardCarrierAgent) ----------------
    path('getTransInwardCarrierAgentTableInfo/', views.GetTransInwardCarrierAgent.as_view()),
    path('getTransInwardCarrierAgent/<str:code>/', views.GetTransInwardCarrierAgentByCode.as_view()),
    path('deleteTransInwardCarrierAgent/<str:code>/', views.DeleteTransInwardCarrierAgent.as_view()),
    path('postTransInwardCarrierAgentTable/', views.PostTransInwardCarrierAgentTable.as_view()),
    path('editTransInwardCarrierAgent/<str:code>/', views.EditTransInwardCarrierAgentByCode.as_view()),
    # ---------------- Outward Carrier Agent (OutwardCarrierAgent) ----------------
    path('getTransOutwardCarrierAgentTableInfo/', views.GetTransOutwardCarrierAgent.as_view()),
    path('getTransOutwardCarrierAgent/<str:code>/', views.GetTransOutwardCarrierAgentByCode.as_view()),
    path('deleteTransOutwardCarrierAgent/<str:code>/', views.DeleteTransOutwardCarrierAgent.as_view()),
    path('postTransOutwardCarrierAgentTable/', views.PostTransOutwardCarrierAgentTable.as_view()),
    path('editTransOutwardCarrierAgent/<str:code>/', views.EditTransOutwardCarrierAgentByCode.as_view()),
    # ---------------- Consignee (OutConsignee) ----------------
    path('getTransConsigneeTableInfo/', views.GetTransConsigneeTable.as_view()),
    path('getTransConsignee/<str:consigneecode>/', views.GetTransConsigneeByCode.as_view()),
    path('deleteTransConsignee/<str:consigneecode>/', views.DeleteTransConsignee.as_view()),
    path('postTransConsigneeTable/', views.PostTransConsigneeTable.as_view()),
    path('editTransConsignee/<str:consigneecode>/', views.EditTransConsigneeByCode.as_view()),
    #EndUser
    path('getTransEndUserTableInfo/', views.GetTransEndUserTable.as_view()),
    path('getTransEndUserByCode/<str:EndUserCode>/', views.GetTransEndUserByCode.as_view()),
    path("deleteTransEndUserByCode/<str:EndUserCode>/", views.DeleteTransEndUserByCode.as_view()),
    path("postTransEndUserTable/", views.PostTransEndUserTable.as_view()),
    path('editTransEndUserByCode/<str:EndUserCode>/', views.EditTransEndUserByCode.as_view()),
#      #Manufacturer
#     path('getOutManufacturerTableInfo/', views.GetOutManufacturerTable.as_view()),
#     path('getOutManufacturer/<str:ManufacturerCode>/', views.GetOutManufacturerByCode.as_view()),
#     path('deleteOutManufacturer/<str:ManufacturerCode>/', views.DeleteOutManufacturerByCode.as_view()),
#     path('postOutManufacturerTable/', views.PostOutManufacturerTable.as_view()),
#     path('editOutManufacturer/<str:ManufacturerCode>/', views.EditOutManufacturerByCode.as_view()),# #     #SupplierManufacturerParty
# # #     path('getInSupplierManufacturerPartyTableInfo/', views.GetInSupplierManufacturerParty.as_view()),
# # #     path('getInSupplierManufacturerPartyByCode/<str:code>/', views.GetInSupplierManufacturerPartyByCode.as_view()),
# # #     path('deleteInSupplierManufacturerParty/<str:code>/', views.DeleteInSupplierManufacturerParty.as_view()),
# # #     path('postInSupplierManufacturerPartyTable/', views.PostInSupplierManufacturerParty.as_view()),
# # #     path('editInSupplierManufacturerPartyByCode/<str:code>/', views.EditInSupplierManufacturerPartyByCode.as_view()),
#TransFile
        path('getTransFileTableInfo/', views.GetTransFileTable.as_view()),
        path("getTransFileByEditPermitId/", views.GetTransFileByEditPermitId.as_view()),
        path('getTransFileByPermitId/<str:permit_id>/', views.GetTransFileByPermitId.as_view()),
        path("deleteTransFile/<str:permit_id>/<int:sno>/", views.DeleteTransFile.as_view()),
        path("postTransFileTable/", views.PostTransFileTable.as_view()),
        path('editTransFileByPermitId/<str:permit_id>/', views.EditTransFileByPermitId.as_view()),
# # #     #InPMT
# # #     path('getInPMTTableInfo/', views.GetInPMTTable.as_view()),
# # #     path('getInPMTByPermitNo/<str:permit_number>/', views.GetInPMTByPermitNo.as_view()),
# # #     path("deleteInPMT/<str:permit_number>/", views.DeleteInPMT.as_view()),
# # #     path("postInPMTTable/", views.PostInPMTTable.as_view()),
# # #     path('editInPMTByPermitNo/<str:permit_number>/', views.EditInPMTByPermitNo.as_view()),
# # #     #InAMDPMT
# # #     path('getInAMDPMTTableInfo/', views.GetInAMDPMTTable.as_view()),
# # #     path('getInAMDPMTByPermitNo/<str:permit_number>/', views.GetInAMDPMTByPermitNo.as_view()),
# # #     path("deleteInAMDPMT/<str:permit_number>/", views.DeleteInAMDPMT.as_view()),
# # #     path("postInAMDPMTTable/", views.PostInAMDPMTTable.as_view()),
# # #     path('editInAMDPMTByPermitNo/<str:permit_number>/', views.EditInAMDPMTByPermitNo.as_view()),
# # #     #InRejectStatus
# # #     path('getInRejectStatusTableInfo/', views.GetInRejectStatusTable.as_view()),
# # #     path('getInRejectStatusByMsgId/<str:msgId>/', views.GetInRejectStatusByMsgId.as_view()),
# # #     path("deleteInRejectStatusByMsgId/<str:msgId>/", views.DeleteInRejectStatusByMsgId.as_view()),
# # #     path("postInRejectStatusTable/", views.PostInRejectStatusTable.as_view()),
# # #     path('editInRejectStatusByMsgId/<str:msgId>/', views.EditInRejectStatusByMsgId.as_view()),
# # #     #InErrorStatus
# # #     path('getInErrorStatusTableInfo/', views.GetInErrorStatusTable.as_view()),
# # #     path('getInErrorStatusByMsgId/<str:msgId>/', views.GetInErrorStatusByMsgId.as_view()),
# # #     path("deleteInErrorStatusByMsgId/<str:msgId>/", views.DeleteInErrorStatusByMsgId.as_view()),
# # #     path("postInErrorStatusTable/", views.PostInErrorStatusTable.as_view()),
# # #     path('editInErrorStatusByMsgId/<str:msgId>/', views.EditInErrorStatusByMsgId.as_view()),
# # # # # -----------------------------------------------------------------------------------------------------------------#
# # #     # Excel Templatedownload
    path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
# #     #Excel Item Upload
    path('uploadedTransItemExcel/', views.TransItemExcelUpload.as_view()),
# #     #Excel Item Upload
    path('editTransAllItems/', views.TransAllItemUpdate.as_view()),
# #     #Delete Hawb by PermitId
    path('deleteTransHawbByPermitId/<str:permit_id>/', views.DeleteTransHawbl.as_view()),
#     # Inpaymentnew


# # #   # Excel Templatedownload
# # #     path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
# # #     #Excel Item Upload
# # #     path('uploadedExcelItem/', views.ItemExcelUpload.as_view()),
# # #     #Excel Item Upload
# # #     path('editAllItems/', views.AllItemUpdate.as_view()),
# # #     #Delete Hawb by PermitId
# # #     path('deleteHawbByPermitId/<str:permit_id>/', views.DeleteHawbl.as_view()),
# # #     # Inpaymentnew
# # #     path('inpaymentnew/',views.InpaymentNewPermit.as_view()),
# # #     path('inpaymentList/',views.InpaymentList.as_view()),
# # #     # PrintGst
# # #     path('PrintGst/<str:PermitId>/', views.PrintGst.as_view()),
# # #     # PrintGst
# # #     path('printGstAll/', views.PrintGstAll.as_view()),
# # #     # Download ccp
# # #     path('downloadCcp/', views.DownloadCcp.as_view()),
# # #     # PRINT CCP
# # #     path('printCcp/<str:permit_id>/',views.PrintCcp.as_view()),
# # #     # Gst status
# #      path('gstInnonStatus/',views.GstInnonStatus.as_view()),
# # #     # Download Data
# # #     path('downloadData/', views.DownloadData.as_view()),
# # #     # Print Gst Via Excel
# # #     path('gstExcel/', views.GstExcel.as_view()),
# # #     # Print Status
# # #     path('printStatus/<str:PermitId>/', views.PrintStatus.as_view()),
# # #     # Xml Submit 
# # #     path('XmlSubmit/', views.XmlSubmit.as_view()),
# # #     # Amend
# # #     path('getAmendPermitByMsgId/', views.GetAmendByMsgId.as_view()),
# # #     path('postAmendTable/', views.PostAmendTable.as_view()),
# # #     # Cancel
# # #     path('getCancelPermitByMsgId/', views.CancelPermit.as_view()),
# # #     path('postCancelPermit/', views.PostCancelPermit.as_view()),
# # #     # Refund
# # #     path('getRefundPermitByMsgId/', views.RefundPermit.as_view()),
# # #     path('postRefundPermit/', views.PostRefundPermit.as_view()),
# # #     path('getRefundValSummaryByMsgId/', views.RefundValSummary.as_view()),
# # #     path('postRefundValSummary/', views.PostRefundValSummary.as_view()),
# # #     path('getRefundItemSummaryByMsgId/', views.GetReundItemSummByMsgId.as_view()),
# # #     path('postRefundItemSummary/', views.PostReundItemSumm.as_view()),
# # #     # get permit conditions for view 
# # #     path('getPermitConditions/', views.GetPermitConditions.as_view()),
# # #     # transmit innonpayment
# # #     path("transmitInnonpayment/", views.TransmitInnonpayment.as_view()),
# # #     # mailbox transmit
# # #     path("mailboxTransmitData/", views.MailBoxTransmitData.as_view()),
# # #     # Get Exchange Rate using Date
# # #     path("getExchangeRateByDate/",views.GetExchangeRateByDate.as_view()),

# Item Copy common

path('syncTransItemFromCommon/', views.SyncTransItemFromCommon.as_view()),

# PostItemwithcasc

path('postTransItemWithCasc/', views.PostTransItemWithCascTable.as_view()),

]



