from django.urls import path
from . import *

urlpatterns = [
 # Header Table
      path('getOutHeaderTableInfo/', views.GetOutHeaderTable.as_view()),
      path('getOutHeaderByPermitId/', views.GetOutHeaderByPermitId.as_view()),
      path('postOutHeaderTable/', views.PostOutHeaderTable.as_view()),
      path("editOutPermit/<str:permit_id>/", views.EditOutHeaderByPermit.as_view()),
      path("deleteOutPermit/", views.DeleteOutPermit.as_view()),

 #CommonInvoiceTable,
    path('outInvoiceTableInfo/', views.OutInvoiceTable.as_view()),
    path("getOutInvoiceNo/<str:invoice_no>/<str:permit_id>/",views.GetOutInvoiceByInvoiceNo.as_view()),
    path("getOutInvoiceByEditPermitId/",views.GetOutInvoiceByEditPermitId.as_view()),
    path("getOutInvoiceByPermitId/<str:permit_id>/",views.GetOutInvoiceByPermitId.as_view()),
    path("deleteOutInvoiceNo/",views.DeleteOutInnvoice.as_view()), 
    path('postOutInvoiceTable/', views.PostOutInvoiceTable.as_view()),
    path('editOutInvoice/<str:invoice_no>/<str:permit_id>/', views.EditOutInvoiceByInvoiceNo.as_view()),


    #CommonItemTable
    path('getOutItemTableInfo/', views.GetOutItemTabel.as_view()),
    path('getOutItemNo/<str:item_no>/<str:permit_id>/', views.GetOutItemByItemNo.as_view()),
    path("getOutItemByEditPermitId/",views.GetOutItemByEditPermitId.as_view()),
    path("deleteOutItem/",views.DeleteOutItem.as_view()),
    path('postOutItemTable/', views.PostOutItemTable.as_view()),
    path('editOutItem/<str:item_no>/<str:permit_id>/', views.EditOutItemByItemNo.as_view()),
 

# # #   # CopyPermit
# # #     path("copyInpayment/", views.CopyInpayment.as_view()),

    #casc
    path('getOutCascTableInfo/', views.GetOutCascTabel.as_view()),
    path('getOutCasc/<str:permit_id>/', views.GetOutCascByPermitId.as_view()),
    path('deleteOutCasc/', views.DeleteOutCasc.as_view()),
    path('deleteOutCascByCascId/<str:casc_id>/<int:row_no>/<str:permit_id>/',views.DeleteOutCascByCascId.as_view()),
    path("postOutCascTable/", views.PostOutCascTable.as_view()),
    path('editOutCasc/<str:permit_id>/', views.EditOutCascByPermitId.as_view()),
#     #cpc
    path('getOutCpcTableInfo/', views.GetOutCpcTabel.as_view()),
    path("getOutCpcByEditPermitId/",views.GetOutCpcByEditPermitId.as_view()),
    path('getOutCpc/<str:permit_id>/', views.GetOutCpcByPermitId.as_view()),
    path("deleteOutCpc/<str:permit_id>/", views.DeleteOutCpc.as_view()),
    path("postOutCpcTable/", views.PostOutCpcTable.as_view()),
    path('editOutCpc/<str:permit_id>/', views.EditOutCpcByPermitId.as_view()),
# #     #container
    path('getOutContainerTableInfo/', views.GetOutContainerTabel.as_view()),
    path("getOutContainerByEditPermitId/",views.GetOutContainerByEditPermitId.as_view()),
    path('getOutContainer/<str:permit_id>/', views.GetOutContainerByPermitId.as_view()),
    path("deleteOutContainer/", views.DeleteOutContainer.as_view()),
    path("postOutContainerTable/", views.PostOutContainerTable.as_view()),
    path('editOutContainer/<str:permit_id>/', views.EditOutContainerByPermitId.as_view()),
# #     #housecode
# #     path('getInHouseItemCode/',views.GetInHouseItemCode.as_view()),
# #     path('getInHouseItemCodeByHouseCode/<str:housecode>/',views.GetInHouseItemCodeByHouseCode.as_view()),
# #     path('postInHouseItemCode/',views.PostInHouseItemCode.as_view()),
# #     path('editInHouseItemCode/<str:housecode>/',views.EditInHouseItemCode.as_view()),
# #     path('deleteInHouseItemCode/<str:housecode>/',views.DeleteInHouseItemCode.as_view()),
    #importer
    path('getOutImporterTableInfo/', views.GetOutImporterTable.as_view()),
    path('getOutImporter/<str:code>/', views.GetOutImporterByCode.as_view()),
    path('deleteOutImporter/<str:code>/', views.DeleteOutImporter.as_view()),
    path('postOutImporterTable/', views.PostOutImporterTable.as_view()),
    path('editOutImporter/<str:code>/', views.EditOutImporterByCode.as_view()),

# # #  HandlingAgent
# #     path('getInHandlingAgentTableInfo/', views.GetInHandlingAgentTabel.as_view()),
# #     path('getInHandlingAgent/<str:code>/', views.GetInHandlingAgentByCode.as_view()),
# #     path("deleteInHandlingAgent/<str:code>/", views.DeleteInHandlingAgent.as_view()),
# #     path("postInHandlingAgentTable/", views.PostInHandlingAgentTable.as_view()),
# #     path('editInHandlingAgent/<str:code>/', views.EditInHandlingAgentByCode.as_view()),
 
    # ---------------- Freight Forwarder (FreightForwarder) ----------------
    path('getOutFreightForwarderTableInfo/', views.GetOutFreightForwarderTable.as_view()),
    path('getOutFreightForwarder/<str:code>/', views.GetOutFreightForwarderByCode.as_view()),
    path('deleteOutFreightForwarder/<str:code>/', views.DeleteOutFreightForwarder.as_view()),
    path('postOutFreightForwarderTable/', views.PostOutFreightForwarderTable.as_view()),
    path('editOutFreightForwarder/<str:code>/', views.EditOutFreightForwarderByCode.as_view()),


#     #ClaimantParty
#     path("getInnonClaimantPartyTable/", views.GetInnonClaimantPartyTable.as_view()),
#     path("getInnonClaimantParty/<int:id>/", views.GetInnonClaimantPartyById.as_view()),
#     path("deleteInnonClaimantParty/<int:id>/", views.DeleteInnonClaimantParty.as_view()),
#     path("postInnonClaimantPartyTable/", views.PostInnonClaimantPartyTable.as_view()),
#     path("editInnonClaimantParty/<int:id>/", views.EditInnonClaimantPartyById.as_view()), 
   
    #Exporter
    path('getOutExporterTableInfo/', views.GetOutExporterTable.as_view()),
    path('getOutExporter/<str:code>/', views.GetOutExporterByCode.as_view()),
    path("deleteOutExporter/<str:code>/", views.DeleteOutExporter.as_view()),
    path("postOutExporterTable/", views.PostOutExporterTable.as_view()),
    path('editOutExporter/<str:code>/', views.EditOutExporterByCode.as_view()),

# ---------------- Inward Carrier Agent (InwardCarrierAgent) ----------------
    path('getOutInwardCarrierAgentTableInfo/', views.GetOutInwardCarrierAgent.as_view()),
    path('getOutInwardCarrierAgent/<str:code>/', views.GetOutInwardCarrierAgentByCode.as_view()),
    path('deleteOutInwardCarrierAgent/<str:code>/', views.DeleteOutInwardCarrierAgent.as_view()),
    path('postOutInwardCarrierAgentTable/', views.PostOutInwardCarrierAgentTable.as_view()),
    path('editOutInwardCarrierAgent/<str:code>/', views.EditOutInwardCarrierAgentByCode.as_view()),
    # ---------------- Outward Carrier Agent (OutwardCarrierAgent) ----------------
    path('getOutOutwardCarrierAgentTableInfo/', views.GetOutOutwardCarrierAgent.as_view()),
    path('getOutOutwardCarrierAgent/<str:code>/', views.GetOutOutwardCarrierAgentByCode.as_view()),
    path('deleteOutOutwardCarrierAgent/<str:code>/', views.DeleteOutOutwardCarrierAgent.as_view()),
    path('postOutOutwardCarrierAgentTable/', views.PostOutOutwardCarrierAgentTable.as_view()),
    path('editOutOutwardCarrierAgent/<str:code>/', views.EditOutOutwardCarrierAgentByCode.as_view()),
    # ---------------- Consignee (OutConsignee) ----------------
    path('getOutConsigneeTableInfo/', views.GetOutConsigneeTable.as_view()),
    path('getOutConsignee/<str:consigneecode>/', views.GetOutConsigneeByCode.as_view()),
    path('deleteOutConsignee/<str:consigneecode>/', views.DeleteOutConsignee.as_view()),
    path('postOutConsigneeTable/', views.PostOutConsigneeTable.as_view()),
    path('editOutConsignee/<str:consigneecode>/', views.EditOutConsigneeByCode.as_view()),
# #     #EndUser
# #     path('getInEndUserTableInfo/', views.GetInEndUserTable.as_view()),
# #     path('getInEndUserByCode/<str:EndUserCode>/', views.GetInEndUserByCode.as_view()),
# #     path("deleteInEndUserByCode/<str:EndUserCode>/", views.DeleteInEndUserByCode.as_view()),
# #     path("postInEndUserTable/", views.PostInEndUserTable.as_view()),
# #     path('editInEndUserByCode/<str:EndUserCode>/', views.EditInEndUserByCode.as_view()),
     #Manufacturer
    path('getOutManufacturerTableInfo/', views.GetOutManufacturerTable.as_view()),
    path('getOutManufacturer/<str:ManufacturerCode>/', views.GetOutManufacturerByCode.as_view()),
    path('deleteOutManufacturer/<str:ManufacturerCode>/', views.DeleteOutManufacturerByCode.as_view()),
    path('postOutManufacturerTable/', views.PostOutManufacturerTable.as_view()),
    path('editOutManufacturer/<str:ManufacturerCode>/', views.EditOutManufacturerByCode.as_view()),# #     #SupplierManufacturerParty
# #     path('getInSupplierManufacturerPartyTableInfo/', views.GetInSupplierManufacturerParty.as_view()),
# #     path('getInSupplierManufacturerPartyByCode/<str:code>/', views.GetInSupplierManufacturerPartyByCode.as_view()),
# #     path('deleteInSupplierManufacturerParty/<str:code>/', views.DeleteInSupplierManufacturerParty.as_view()),
# #     path('postInSupplierManufacturerPartyTable/', views.PostInSupplierManufacturerParty.as_view()),
# #     path('editInSupplierManufacturerPartyByCode/<str:code>/', views.EditInSupplierManufacturerPartyByCode.as_view()),
    #OutFile
    path('getOutFileTableInfo/', views.GetOutFileTable.as_view()),
    path("getOutFileByEditPermitId/", views.GetOutFileByEditPermitId.as_view()),
    path('getOutFileByPermitId/<str:permit_id>/', views.GetOutFileByPermitId.as_view()),
    path("deleteOutFile/<str:permit_id>/<int:sno>/", views.DeleteOutFile.as_view()),
    path("postOutFileTable/", views.PostOutFileTable.as_view()),
    path('editOutFileByPermitId/<str:permit_id>/', views.EditOutFileByPermitId.as_view()),
# #     #InPMT
# #     path('getInPMTTableInfo/', views.GetInPMTTable.as_view()),
# #     path('getInPMTByPermitNo/<str:permit_number>/', views.GetInPMTByPermitNo.as_view()),
# #     path("deleteInPMT/<str:permit_number>/", views.DeleteInPMT.as_view()),
# #     path("postInPMTTable/", views.PostInPMTTable.as_view()),
# #     path('editInPMTByPermitNo/<str:permit_number>/', views.EditInPMTByPermitNo.as_view()),
# #     #InAMDPMT
# #     path('getInAMDPMTTableInfo/', views.GetInAMDPMTTable.as_view()),
# #     path('getInAMDPMTByPermitNo/<str:permit_number>/', views.GetInAMDPMTByPermitNo.as_view()),
# #     path("deleteInAMDPMT/<str:permit_number>/", views.DeleteInAMDPMT.as_view()),
# #     path("postInAMDPMTTable/", views.PostInAMDPMTTable.as_view()),
# #     path('editInAMDPMTByPermitNo/<str:permit_number>/', views.EditInAMDPMTByPermitNo.as_view()),
# #     #InRejectStatus
# #     path('getInRejectStatusTableInfo/', views.GetInRejectStatusTable.as_view()),
# #     path('getInRejectStatusByMsgId/<str:msgId>/', views.GetInRejectStatusByMsgId.as_view()),
# #     path("deleteInRejectStatusByMsgId/<str:msgId>/", views.DeleteInRejectStatusByMsgId.as_view()),
# #     path("postInRejectStatusTable/", views.PostInRejectStatusTable.as_view()),
# #     path('editInRejectStatusByMsgId/<str:msgId>/', views.EditInRejectStatusByMsgId.as_view()),
# #     #InErrorStatus
# #     path('getInErrorStatusTableInfo/', views.GetInErrorStatusTable.as_view()),
# #     path('getInErrorStatusByMsgId/<str:msgId>/', views.GetInErrorStatusByMsgId.as_view()),
# #     path("deleteInErrorStatusByMsgId/<str:msgId>/", views.DeleteInErrorStatusByMsgId.as_view()),
# #     path("postInErrorStatusTable/", views.PostInErrorStatusTable.as_view()),
# #     path('editInErrorStatusByMsgId/<str:msgId>/', views.EditInErrorStatusByMsgId.as_view()),
# # # # -----------------------------------------------------------------------------------------------------------------#
# # #     # Excel Templatedownload
# #     path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
# #     #Excel Item Upload
    path('uploadedOutItemExcel/', views.OutItemExcelUpload.as_view()),
# #     #Excel Item Upload
    path('editOutAllItems/', views.OutAllItemUpdate.as_view()),
# #     #Delete Hawb by PermitId
    path('deleteOutHawbByPermitId/<str:permit_id>/', views.DeleteOutHawbl.as_view()),
#     # Inpaymentnew


# #   # Excel Templatedownload
# #     path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
# #     #Excel Item Upload
# #     path('uploadedExcelItem/', views.ItemExcelUpload.as_view()),
# #     #Excel Item Upload
# #     path('editAllItems/', views.AllItemUpdate.as_view()),
# #     #Delete Hawb by PermitId
# #     path('deleteHawbByPermitId/<str:permit_id>/', views.DeleteHawbl.as_view()),
# #     # Inpaymentnew
# #     path('inpaymentnew/',views.InpaymentNewPermit.as_view()),
# #     path('inpaymentList/',views.InpaymentList.as_view()),
# #     # PrintGst
# #     path('PrintGst/<str:PermitId>/', views.PrintGst.as_view()),
# #     # PrintGst
# #     path('printGstAll/', views.PrintGstAll.as_view()),
# #     # Download ccp
# #     path('downloadCcp/', views.DownloadCcp.as_view()),
# #     # PRINT CCP
# #     path('printCcp/<str:permit_id>/',views.PrintCcp.as_view()),
# #     # Gst status
#      path('gstInnonStatus/',views.GstInnonStatus.as_view()),
# #     # Download Data
# #     path('downloadData/', views.DownloadData.as_view()),
# #     # Print Gst Via Excel
# #     path('gstExcel/', views.GstExcel.as_view()),
# #     # Print Status
# #     path('printStatus/<str:PermitId>/', views.PrintStatus.as_view()),
# #     # Xml Submit 
# #     path('XmlSubmit/', views.XmlSubmit.as_view()),
# #     # Amend
# #     path('getAmendPermitByMsgId/', views.GetAmendByMsgId.as_view()),
# #     path('postAmendTable/', views.PostAmendTable.as_view()),
# #     # Cancel
# #     path('getCancelPermitByMsgId/', views.CancelPermit.as_view()),
# #     path('postCancelPermit/', views.PostCancelPermit.as_view()),
# #     # Refund
# #     path('getRefundPermitByMsgId/', views.RefundPermit.as_view()),
# #     path('postRefundPermit/', views.PostRefundPermit.as_view()),
# #     path('getRefundValSummaryByMsgId/', views.RefundValSummary.as_view()),
# #     path('postRefundValSummary/', views.PostRefundValSummary.as_view()),
# #     path('getRefundItemSummaryByMsgId/', views.GetReundItemSummByMsgId.as_view()),
# #     path('postRefundItemSummary/', views.PostReundItemSumm.as_view()),
# #     # get permit conditions for view 
# #     path('getPermitConditions/', views.GetPermitConditions.as_view()),
# #     # transmit innonpayment
# #     path("transmitInnonpayment/", views.TransmitInnonpayment.as_view()),
# #     # mailbox transmit
# #     path("mailboxTransmitData/", views.MailBoxTransmitData.as_view()),
# #     # Get Exchange Rate using Date
# #     path("getExchangeRateByDate/",views.GetExchangeRateByDate.as_view()),
path('syncOutItemFromCommon/',views.SyncOutItemFromCommon.as_view()),

# postitemwithcasc

path('postOutItemWithCasc/', views.PostOutItemWithCascTable.as_view()),
]



