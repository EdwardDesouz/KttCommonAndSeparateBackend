from django.urls import path
from . import *

urlpatterns = [
 # Header Table
      path('getInnonHeaderTableInfo/', views.GetInnonHeaderTable.as_view()),
      path('getInnonHeaderByPermitId/', views.GetInnonHeaderByPermitId.as_view()),
      path('postInnonHeaderTable/', views.PostInnonHeaderTable.as_view()),
      path("editInnonPermit/<str:permit_id>/", views.EditInnonHeaderByPermit.as_view()),
      path("deleteInnonPermit/", views.DeleteInnonPermit.as_view()),

 #CommonInvoiceTable,
    path('innonInvoiceTableInfo/', views.InnonInvoiceTable.as_view()),
    path("getInnonInvoiceNo/<str:invoice_no>/<str:permit_id>/",views.GetInnonInvoiceByInvoiceNo.as_view()),
    path("getInnonInvoiceByEditPermitId/",views.GetInnonInvoiceByEditPermitId.as_view()),
    path("getInnonInvoiceByPermitId/<str:permit_id>/",views.GetInnonInvoiceByPermitId.as_view()),
    path("deleteInnonInvoiceNo/",views.DeleteInnonInnvoice.as_view()), 
    path('postInnonInvoiceTable/', views.PostInnonInvoiceTable.as_view()),
    path('editInnonInvoice/<str:invoice_no>/<str:permit_id>/', views.EditInnonInvoiceByInvoiceNo.as_view()),


    #CommonItemTable
    path('getInnonItemTableInfo/', views.GetInnonItemTabel.as_view()),
    path('getInnonItemNo/<str:item_no>/<str:permit_id>/', views.GetInnonItemByItemNo.as_view()),
    path("getInnonItemByEditPermitId/",views.GetInnonItemByEditPermitId.as_view()),
    path("deleteInnonItem/",views.DeleteInnonItem.as_view()),
    path('postInnonItemTable/', views.PostInnonItemTable.as_view()),
    path('editInnonItem/<str:item_no>/<str:permit_id>/', views.EditInnonItemByItemNo.as_view()),
 

# #   # CopyPermit
# #     path("copyInpayment/", views.CopyInpayment.as_view()),

    #casc
    path('getInnonCascTableInfo/', views.GetInnonCascTabel.as_view()),
    path('getInnonCasc/<str:permit_id>/', views.GetInnonCascByPermitId.as_view()),
    path('deleteInnonCasc/', views.DeleteInnonCasc.as_view()),
    path('deleteInnonCascByCascId/<str:casc_id>/<int:row_no>/<str:permit_id>/',views.DeleteInnonCascByCascId.as_view()),
    path("postInnonCascTable/", views.PostInnonCascTable.as_view()),
    path('editInnonCasc/<str:permit_id>/', views.EditInnonCascByPermitId.as_view()),
#     #cpc
    path('getInnonCpcTableInfo/', views.GetInnonCpcTabel.as_view()),
    path("getInnonCpcByEditPermitId/",views.GetInnonCpcByEditPermitId.as_view()),
    path('getInnonCpc/<str:permit_id>/', views.GetInnonCpcByPermitId.as_view()),
    path("deleteInnonCpc/<str:permit_id>/", views.DeleteInnonCpc.as_view()),
    path("postInnonCpcTable/", views.PostInnonCpcTable.as_view()),
    path('editInnonCpc/<str:permit_id>/', views.EditInnonCpcByPermitId.as_view()),
#     #container
    path('getInnonContainerTableInfo/', views.GetInnonContainerTabel.as_view()),
    path("getInnonContainerByEditPermitId/",views.GetInnonContainerByEditPermitId.as_view()),
    path('getInnonContainer/<str:permit_id>/', views.GetInnonContainerByPermitId.as_view()),
    path("deleteInnonContainer/", views.DeleteInnonContainer.as_view()),
    path("postInnonContainerTable/", views.PostInnonContainerTable.as_view()),
    path('editInnonContainer/<str:permit_id>/', views.EditInnonContainerByPermitId.as_view()),
#     #housecode
#     path('getInHouseItemCode/',views.GetInHouseItemCode.as_view()),
#     path('getInHouseItemCodeByHouseCode/<str:housecode>/',views.GetInHouseItemCodeByHouseCode.as_view()),
#     path('postInHouseItemCode/',views.PostInHouseItemCode.as_view()),
#     path('editInHouseItemCode/<str:housecode>/',views.EditInHouseItemCode.as_view()),
#     path('deleteInHouseItemCode/<str:housecode>/',views.DeleteInHouseItemCode.as_view()),
    #importer
    path('getInnonImporterTableInfo/', views.GetInnonImporterTable.as_view()),
    path('getInnonImporter/<str:code>/', views.GetInnonImporterByCode.as_view()),
    path("deleteInnonImporter/<str:code>/", views.DeleteInnonImporter.as_view()),
    path("postInnonImporterTable/", views.PostInnonImporterTable.as_view()),
    path('editInnonImporter/<str:code>/', views.EditInnonImporterByCode.as_view()),
# #  HandlingAgent
#     path('getInHandlingAgentTableInfo/', views.GetInHandlingAgentTabel.as_view()),
#     path('getInHandlingAgent/<str:code>/', views.GetInHandlingAgentByCode.as_view()),
#     path("deleteInHandlingAgent/<str:code>/", views.DeleteInHandlingAgent.as_view()),
#     path("postInHandlingAgentTable/", views.PostInHandlingAgentTable.as_view()),
#     path('editInHandlingAgent/<str:code>/', views.EditInHandlingAgentByCode.as_view()),
    #FreightForwarder
    path("getInnonFreightForwarderTable/",views.GetInnonFreightForwarderTable.as_view()),
    path("getInnonFreightForwarder/<str:code>/",views.GetInnonFreightForwarderByCode.as_view()),
    path("deleteInnonFreightForwarder/<str:code>/",views.DeleteInnonFreightForwarder.as_view()),
    path("postInnonFreightForwarderTable/",views.PostInnonFreightForwarderTable.as_view()),
    path("editInnonFreightForwarder/<str:code>/",views.EditInnonFreightForwarderByCode.as_view()),
    #ClaimantParty
    path("getInnonClaimantPartyTable/", views.GetInnonClaimantPartyTable.as_view()),
    path("getInnonClaimantParty/<int:id>/", views.GetInnonClaimantPartyById.as_view()),
    path("deleteInnonClaimantParty/<int:id>/", views.DeleteInnonClaimantParty.as_view()),
    path("postInnonClaimantPartyTable/", views.PostInnonClaimantPartyTable.as_view()),
    path("editInnonClaimantParty/<int:id>/", views.EditInnonClaimantPartyById.as_view()), 
   
    #Exporter
    path('getInnonExporterTableInfo/', views.GetInnonExporterTabel.as_view()),
    path('getInnonExporter/<str:code>/', views.GetInnonExporterByCode.as_view()),
    path("deleteInnonExporter/<str:code>/", views.DeleteInnonExporter.as_view()),
    path("postInnonExporterTable/", views.PostInnonExporterTable.as_view()),
    path('editInnonExporter/<str:code>/', views.EditInnonExporterByCode.as_view()),
    #Inward
    path('getInnonInwardCarrierAgentTableInfo/', views.GetInnonInwardCarrierAgent.as_view()),
    path('getInnonInwardCarrierAgent/<str:code>/', views.GetInnonInwardCarrierAgentByCode.as_view()),
    path("deleteInnonInwardCarrierAgent/<str:code>/", views.DeleteInnonInwardCarrierAgent.as_view()),
    path("postInnonInwardCarrierAgentTable/", views.PostInnonInwardCarrierAgentTable.as_view()),
    path('editInnonInwardCarrierAgent/<str:code>/', views.EditInnonInwardCarrierAgentByCode.as_view()),
#Outward
    path('getInnonOutwardCarrierAgentTableInfo/', views.GetInnonOutwardCarrierAgent.as_view()),
    path('getInnonOutwardCarrierAgentByCode/<str:code>/', views.GetInnonOutwardCarrierAgentByCode.as_view()),
    path("deleteInnonOutwardCarrierAgent/<str:code>/", views.DeleteInnonOutwardCarrierAgent.as_view()),
    path("postInnonOutwardCarrierAgentTable/", views.PostInnonOutwardCarrierAgentTable.as_view()),
    path('editInnonOutwardCarrierAgentByCode/<str:code>/', views.EditInnonOutwardCarrierAgentByCode.as_view()),

#Consignee
    path('getInnonConsigneeTableInfo/', views.GetInnonConsigneeTable.as_view()),
    path('getInnonConsigneeByCode/<str:consigneecode>/', views.GetInnonConsigneeByCode.as_view()),
    path("deleteInnonConsigneeAgent/<str:consigneecode>/", views.DeleteInnonConsigneeAgent.as_view()),
    path("postInnonConsigneeTable/", views.PostInnonConsigneeTable.as_view()),
    path('editInnonConsigneeByCode/<str:consigneecode>/', views.EditInnonConsigneeByCode.as_view()),
#     #EndUser
#     path('getInEndUserTableInfo/', views.GetInEndUserTable.as_view()),
#     path('getInEndUserByCode/<str:EndUserCode>/', views.GetInEndUserByCode.as_view()),
#     path("deleteInEndUserByCode/<str:EndUserCode>/", views.DeleteInEndUserByCode.as_view()),
#     path("postInEndUserTable/", views.PostInEndUserTable.as_view()),
#     path('editInEndUserByCode/<str:EndUserCode>/', views.EditInEndUserByCode.as_view()),
#     #Manufacturer
#     path('getInManufacturerTableInfo/', views.GetInManufacturerTable.as_view()),
#     path('getInManufacturerByCode/<str:ManufacturerCode>/', views.GetInManufacturerByCode.as_view()),
#     path('deleteInManufacturerByCode/<str:ManufacturerCode>/', views.DeleteInManufacturerByCode.as_view()),
#     path('postInManufacturerTable/', views.PostInManufacturerTable.as_view()),
#     path('editInManufacturerByCode/<str:ManufacturerCode>/', views.EditInManufacturerByCode.as_view()),
#     #SupplierManufacturerParty
#     path('getInSupplierManufacturerPartyTableInfo/', views.GetInSupplierManufacturerParty.as_view()),
#     path('getInSupplierManufacturerPartyByCode/<str:code>/', views.GetInSupplierManufacturerPartyByCode.as_view()),
#     path('deleteInSupplierManufacturerParty/<str:code>/', views.DeleteInSupplierManufacturerParty.as_view()),
#     path('postInSupplierManufacturerPartyTable/', views.PostInSupplierManufacturerParty.as_view()),
#     path('editInSupplierManufacturerPartyByCode/<str:code>/', views.EditInSupplierManufacturerPartyByCode.as_view()),
    #InnonFile
    path('getInnonFileTableInfo/', views.GetInnonFileTable.as_view()),
    path("getInnonFileByEditPermitId/", views.GetInnonFileByEditPermitId.as_view()),
    path('getInnonFileByPermitId/<str:permit_id>/', views.GetInnonFileByPermitId.as_view()),
    path("deleteInnonFile/<str:permit_id>/<int:sno>/", views.DeleteInnonFile.as_view()),
    path("postInnonFileTable/", views.PostInnonFileTable.as_view()),
    path('editInnonFileByPermitId/<str:permit_id>/', views.EditInnonFileByPermitId.as_view()),
#     #InPMT
#     path('getInPMTTableInfo/', views.GetInPMTTable.as_view()),
#     path('getInPMTByPermitNo/<str:permit_number>/', views.GetInPMTByPermitNo.as_view()),
#     path("deleteInPMT/<str:permit_number>/", views.DeleteInPMT.as_view()),
#     path("postInPMTTable/", views.PostInPMTTable.as_view()),
#     path('editInPMTByPermitNo/<str:permit_number>/', views.EditInPMTByPermitNo.as_view()),
#     #InAMDPMT
#     path('getInAMDPMTTableInfo/', views.GetInAMDPMTTable.as_view()),
#     path('getInAMDPMTByPermitNo/<str:permit_number>/', views.GetInAMDPMTByPermitNo.as_view()),
#     path("deleteInAMDPMT/<str:permit_number>/", views.DeleteInAMDPMT.as_view()),
#     path("postInAMDPMTTable/", views.PostInAMDPMTTable.as_view()),
#     path('editInAMDPMTByPermitNo/<str:permit_number>/', views.EditInAMDPMTByPermitNo.as_view()),
#     #InRejectStatus
#     path('getInRejectStatusTableInfo/', views.GetInRejectStatusTable.as_view()),
#     path('getInRejectStatusByMsgId/<str:msgId>/', views.GetInRejectStatusByMsgId.as_view()),
#     path("deleteInRejectStatusByMsgId/<str:msgId>/", views.DeleteInRejectStatusByMsgId.as_view()),
#     path("postInRejectStatusTable/", views.PostInRejectStatusTable.as_view()),
#     path('editInRejectStatusByMsgId/<str:msgId>/', views.EditInRejectStatusByMsgId.as_view()),
#     #InErrorStatus
#     path('getInErrorStatusTableInfo/', views.GetInErrorStatusTable.as_view()),
#     path('getInErrorStatusByMsgId/<str:msgId>/', views.GetInErrorStatusByMsgId.as_view()),
#     path("deleteInErrorStatusByMsgId/<str:msgId>/", views.DeleteInErrorStatusByMsgId.as_view()),
#     path("postInErrorStatusTable/", views.PostInErrorStatusTable.as_view()),
#     path('editInErrorStatusByMsgId/<str:msgId>/', views.EditInErrorStatusByMsgId.as_view()),
# # # -----------------------------------------------------------------------------------------------------------------#
# #     # Excel Templatedownload
#     path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
# #     #Excel Item Upload
    path('uploadedInnonItemExcel/', views.InnonItemExcelUpload.as_view()),
# #     #Excel Item Upload
    path('editInnonAllItems/', views.InnonAllItemUpdate.as_view()),
# #     #Delete Hawb by PermitId
#     path('deleteInnonHawbByPermitId/<str:permit_id>/', views.DeleteInnonHawbl.as_view()),
# #     # Inpaymentnew


#   # Excel Templatedownload
#     path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
#     #Excel Item Upload
#     path('uploadedExcelItem/', views.ItemExcelUpload.as_view()),
#     #Excel Item Upload
#     path('editAllItems/', views.AllItemUpdate.as_view()),
#     #Delete Hawb by PermitId
#     path('deleteHawbByPermitId/<str:permit_id>/', views.DeleteHawbl.as_view()),
#     # Inpaymentnew
#     path('inpaymentnew/',views.InpaymentNewPermit.as_view()),
#     path('inpaymentList/',views.InpaymentList.as_view()),
#     # PrintGst
#     path('PrintGst/<str:PermitId>/', views.PrintGst.as_view()),
#     # PrintGst
#     path('printGstAll/', views.PrintGstAll.as_view()),
#     # Download ccp
#     path('downloadCcp/', views.DownloadCcp.as_view()),
#     # PRINT CCP
#     path('printCcp/<str:permit_id>/',views.PrintCcp.as_view()),
#     # Gst status
     path('gstInnonStatus/',views.GstInnonStatus.as_view()),
#     # Download Data
#     path('downloadData/', views.DownloadData.as_view()),
#     # Print Gst Via Excel
#     path('gstExcel/', views.GstExcel.as_view()),
#     # Print Status
#     path('printStatus/<str:PermitId>/', views.PrintStatus.as_view()),
#     # Xml Submit 
#     path('XmlSubmit/', views.XmlSubmit.as_view()),
#     # Amend
#     path('getAmendPermitByMsgId/', views.GetAmendByMsgId.as_view()),
#     path('postAmendTable/', views.PostAmendTable.as_view()),
#     # Cancel
#     path('getCancelPermitByMsgId/', views.CancelPermit.as_view()),
#     path('postCancelPermit/', views.PostCancelPermit.as_view()),
#     # Refund
#     path('getRefundPermitByMsgId/', views.RefundPermit.as_view()),
#     path('postRefundPermit/', views.PostRefundPermit.as_view()),
#     path('getRefundValSummaryByMsgId/', views.RefundValSummary.as_view()),
#     path('postRefundValSummary/', views.PostRefundValSummary.as_view()),
#     path('getRefundItemSummaryByMsgId/', views.GetReundItemSummByMsgId.as_view()),
#     path('postRefundItemSummary/', views.PostReundItemSumm.as_view()),
#     # get permit conditions for view 
#     path('getPermitConditions/', views.GetPermitConditions.as_view()),
#     # transmit innonpayment
#     path("transmitInnonpayment/", views.TransmitInnonpayment.as_view()),
#     # mailbox transmit
#     path("mailboxTransmitData/", views.MailBoxTransmitData.as_view()),
#     # Get Exchange Rate using Date
#     path("getExchangeRateByDate/",views.GetExchangeRateByDate.as_view()),
# item copy delte,add upload
path('syncInnonItemFromCommon/',views.SyncInnonItemFromCommon.as_view()),

# post innonwith cascitem
path('postInnonItemWithCasc/',views.PostInnonItemWithCascTable.as_view()),

    path('deleteInHawbByPermitId/<str:permit_id>/', views.DeleteInHawbl.as_view()),
]



