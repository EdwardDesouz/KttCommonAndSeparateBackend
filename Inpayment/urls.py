from django.urls import path
from . import *

urlpatterns = [
    # Header Table
    path('getInHeaderTableInfo/', views.GetInHeaderTable.as_view()),
    path('getInHeaderByPermitId/', views.GetInHeaderByPermitId.as_view()),
    path('postInHeaderTable/', views.PostInHeaderTable.as_view()),
    path("editPermit/<str:permit_id>/", views.EditInHeaderByPermit.as_view()),
    path("deleteInPermit/", views.DeleteInPermit.as_view()),

 #CommonInvoiceTable,
    path('inInvoiceTableInfo/', views.InInvoiceTable.as_view()),
    path("getInInvoiceNo/<str:invoice_no>/<str:permit_id>/",views.GetInInvoiceByInvoiceNo.as_view()),
    path("getInInvoiceByEditPermitId/",views.GetInInvoiceByEditPermitId.as_view()),
    path("getInInvoiceByPermitId/<str:permit_id>/",views.GetInInvoiceByPermitId.as_view()),
    path("deleteInInvoiceNo/",views.DeleteInInnvoice.as_view()), 
    path('postInInvoiceTable/', views.PostInInvoiceTable.as_view()),
    path('editInInvoice/<str:invoice_no>/<str:permit_id>/', views.EditInInvoiceByInvoiceNo.as_view()),


    #CommonItemTable
    path('getInItemTableInfo/', views.GetInItemTabel.as_view()),
    path('getInItemNo/<str:item_no>/<str:permit_id>/', views.GetInItemByItemNo.as_view()),
    path("getInItemByEditPermitId/",views.GetInItemByEditPermitId.as_view()),
    path("deleteInItem/",views.DeleteInItem.as_view()),
    path('postInItemTable/', views.PostInItemTable.as_view()),
    path('editInItem/<str:item_no>/<str:permit_id>/', views.EditInItemByItemNo.as_view()),
 

#   # CopyPermit
#     path("copyInpayment/", views.CopyInpayment.as_view()),

    #casc
    path('getInCascTableInfo/', views.GetInCascTabel.as_view()),
    path('getInCasc/<str:permit_id>/', views.GetInCascByPermitId.as_view()),
    path('deleteInCasc/', views.DeleteInCasc.as_view()),
    path('deleteInCascByCascId/<str:casc_id>/<int:row_no>/<str:permit_id>/',views.DeleteInCascByCascId.as_view()),
    path("postInCascTable/", views.PostInCascTable.as_view()),
    path('editInCasc/<str:permit_id>/', views.EditInCascByPermitId.as_view()),
    #cpc
    path('getInCpcTableInfo/', views.GetInCpcTabel.as_view()),
    path("getCpcByEditPermitId/",views.GetInCpcByEditPermitId.as_view()),
    path('getCpc/<str:permit_id>/', views.GetInCpcByPermitId.as_view()),
    path("deleteInCpc/<str:permit_id>/", views.DeleteInCpc.as_view()),
    path("postInCpcTable/", views.PostInCpcTable.as_view()),
    path('editInCpc/<str:permit_id>/', views.EditInCpcByPermitId.as_view()),
#     #container
    path('getInContainerTableInfo/', views.GetInContainerTabel.as_view()),
    path("getInContainerByEditPermitId/",views.GetInContainerByEditPermitId.as_view()),
    path('getInContainer/<str:permit_id>/', views.GetInContainerByPermitId.as_view()),
    path("deleteInContainer/", views.DeleteInContainer.as_view()),
    path("postInContainerTable/", views.PostInContainerTable.as_view()),
    path('editInContainer/<str:permit_id>/', views.EditInContainerByPermitId.as_view()),
    #housecode
    path('getInHouseItemCode/',views.GetInHouseItemCode.as_view()),
    path('getInHouseItemCodeByHouseCode/<str:housecode>/',views.GetInHouseItemCodeByHouseCode.as_view()),
    path('postInHouseItemCode/',views.PostInHouseItemCode.as_view()),
    path('editInHouseItemCode/<str:housecode>/',views.EditInHouseItemCode.as_view()),
    path('deleteInHouseItemCode/<str:housecode>/',views.DeleteInHouseItemCode.as_view()),
    #importer
    path('getInImporterTableInfo/', views.GetInImporterTable.as_view()),
    path('getInImporter/<str:code>/', views.GetInImporterByCode.as_view()),
    path("deleteInImporter/<str:code>/", views.DeleteInImporter.as_view()),
    path("postInImporterTable/", views.PostInImporterTable.as_view()),
    path('editInImporter/<str:code>/', views.EditInImporterByCode.as_view()),
#  HandlingAgent
    path('getInHandlingAgentTableInfo/', views.GetInHandlingAgentTabel.as_view()),
    path('getInHandlingAgent/<str:code>/', views.GetInHandlingAgentByCode.as_view()),
    path("deleteInHandlingAgent/<str:code>/", views.DeleteInHandlingAgent.as_view()),
    path("postInHandlingAgentTable/", views.PostInHandlingAgentTable.as_view()),
    path('editInHandlingAgent/<str:code>/', views.EditInHandlingAgentByCode.as_view()),
    #FreightForwarder
    path("getInFreightForwarderTable/",views.GetInFreightForwarderTable.as_view()),
    path("getInFreightForwarder/<str:code>/",views.GetInFreightForwarderByCode.as_view()),
    path("deleteInFreightForwarder/<str:code>/",views.DeleteInFreightForwarder.as_view()),
    path("postInFreightForwarderTable/",views.PostInFreightForwarderTable.as_view()),
    path("editInFreightForwarder/<str:code>/",views.EditInFreightForwarderByCode.as_view()),
    #ClaimantParty
    path("getInClaimantPartyTable/", views.GetInClaimantPartyTable.as_view()),
    path("getInClaimantParty/<int:id>/", views.GetInClaimantPartyById.as_view()),
    path("deleteInClaimantParty/<int:id>/", views.DeleteInClaimantParty.as_view()),
    path("postInClaimantPartyTable/", views.PostInClaimantPartyTable.as_view()),
    path("editInClaimantParty/<int:id>/", views.EditInClaimantPartyById.as_view()), 
   
    #Exporter
    path('getInExporterTableInfo/', views.GetInExporterTabel.as_view()),
    path('getInExporter/<str:code>/', views.GetInExporterByCode.as_view()),
    path("deleteInExporter/<str:code>/", views.DeleteInExporter.as_view()),
    path("postInExporterTable/", views.PostInExporterTable.as_view()),
    path('editInExporter/<str:code>/', views.EditInExporterByCode.as_view()),
    #Inward
    path('getInInwardCarrierAgentTableInfo/', views.GetInInwardCarrierAgent.as_view()),
    path('getInInwardCarrierAgent/<str:code>/', views.GetInInwardCarrierAgentByCode.as_view()),
    path("deleteInInwardCarrierAgent/<str:code>/", views.DeleteInInwardCarrierAgent.as_view()),
    path("postInInwardCarrierAgentTable/", views.PostInInwardCarrierAgentTable.as_view()),
    path('editInInwardCarrierAgent/<str:code>/', views.EditInInwardCarrierAgentByCode.as_view()),
#Outward
    path('getInOutwardCarrierAgentTableInfo/', views.GetInOutwardCarrierAgent.as_view()),
    path('getInOutwardCarrierAgentByCode/<str:code>/', views.GetInOutwardCarrierAgentByCode.as_view()),
    path("deleteInOutwardCarrierAgent/<str:code>/", views.DeleteInOutwardCarrierAgent.as_view()),
    path("postInOutwardCarrierAgentTable/", views.PostInOutwardCarrierAgentTable.as_view()),
    path('editInOutwardCarrierAgentByCode/<str:code>/', views.EditInOutwardCarrierAgentByCode.as_view()),

#Consignee
    path('getInConsigneeTableInfo/', views.GetInConsigneeTable.as_view()),
    path('getInConsigneeByCode/<str:consigneecode>/', views.GetInConsigneeByCode.as_view()),
    path("deleteInConsigneeAgent/<str:consigneecode>/", views.DeleteInConsigneeAgent.as_view()),
    path("postInConsigneeTable/", views.PostInConsigneeTable.as_view()),
    path('editInConsigneeByCode/<str:consigneecode>/', views.EditInConsigneeByCode.as_view()),
    #EndUser
    path('getInEndUserTableInfo/', views.GetInEndUserTable.as_view()),
    path('getInEndUserByCode/<str:EndUserCode>/', views.GetInEndUserByCode.as_view()),
    path("deleteInEndUserByCode/<str:EndUserCode>/", views.DeleteInEndUserByCode.as_view()),
    path("postInEndUserTable/", views.PostInEndUserTable.as_view()),
    path('editInEndUserByCode/<str:EndUserCode>/', views.EditInEndUserByCode.as_view()),
    #Manufacturer
    path('getInManufacturerTableInfo/', views.GetInManufacturerTable.as_view()),
    path('getInManufacturerByCode/<str:ManufacturerCode>/', views.GetInManufacturerByCode.as_view()),
    path('deleteInManufacturerByCode/<str:ManufacturerCode>/', views.DeleteInManufacturerByCode.as_view()),
    path('postInManufacturerTable/', views.PostInManufacturerTable.as_view()),
    path('editInManufacturerByCode/<str:ManufacturerCode>/', views.EditInManufacturerByCode.as_view()),
    #SupplierManufacturerParty
    path('getInSupplierManufacturerPartyTableInfo/', views.GetInSupplierManufacturerParty.as_view()),
    path('getInSupplierManufacturerPartyByCode/<str:code>/', views.GetInSupplierManufacturerPartyByCode.as_view()),
    path('deleteInSupplierManufacturerParty/<str:code>/', views.DeleteInSupplierManufacturerParty.as_view()),
    path('postInSupplierManufacturerPartyTable/', views.PostInSupplierManufacturerParty.as_view()),
    path('editInSupplierManufacturerPartyByCode/<str:code>/', views.EditInSupplierManufacturerPartyByCode.as_view()),
    #InFile
    path('getInFileTableInfo/', views.GetInFileTable.as_view()),
    path("getInFileByEditPermitId/", views.GetInFileByEditPermitId.as_view()),
    path('getInFileByPermitId/<str:permit_id>/', views.GetInFileByPermitId.as_view()),
    path("deleteInFile/<str:permit_id>/<int:sno>/", views.DeleteInFile.as_view()),
    path("postInFileTable/", views.PostInFileTable.as_view()),
    path('editInFileByPermitId/<str:permit_id>/', views.EditInFileByPermitId.as_view()),
    #InPMT
    path('getInPMTTableInfo/', views.GetInPMTTable.as_view()),
    path('getInPMTByPermitNo/<str:permit_number>/', views.GetInPMTByPermitNo.as_view()),
    path("deleteInPMT/<str:permit_number>/", views.DeleteInPMT.as_view()),
    path("postInPMTTable/", views.PostInPMTTable.as_view()),
    path('editInPMTByPermitNo/<str:permit_number>/', views.EditInPMTByPermitNo.as_view()),
    #InAMDPMT
    path('getInAMDPMTTableInfo/', views.GetInAMDPMTTable.as_view()),
    path('getInAMDPMTByPermitNo/<str:permit_number>/', views.GetInAMDPMTByPermitNo.as_view()),
    path("deleteInAMDPMT/<str:permit_number>/", views.DeleteInAMDPMT.as_view()),
    path("postInAMDPMTTable/", views.PostInAMDPMTTable.as_view()),
    path('editInAMDPMTByPermitNo/<str:permit_number>/', views.EditInAMDPMTByPermitNo.as_view()),
    #InRejectStatus
    path('getInRejectStatusTableInfo/', views.GetInRejectStatusTable.as_view()),
    path('getInRejectStatusByMsgId/<str:msgId>/', views.GetInRejectStatusByMsgId.as_view()),
    path("deleteInRejectStatusByMsgId/<str:msgId>/", views.DeleteInRejectStatusByMsgId.as_view()),
    path("postInRejectStatusTable/", views.PostInRejectStatusTable.as_view()),
    path('editInRejectStatusByMsgId/<str:msgId>/', views.EditInRejectStatusByMsgId.as_view()),
    #InErrorStatus
    path('getInErrorStatusTableInfo/', views.GetInErrorStatusTable.as_view()),
    path('getInErrorStatusByMsgId/<str:msgId>/', views.GetInErrorStatusByMsgId.as_view()),
    path("deleteInErrorStatusByMsgId/<str:msgId>/", views.DeleteInErrorStatusByMsgId.as_view()),
    path("postInErrorStatusTable/", views.PostInErrorStatusTable.as_view()),
    path('editInErrorStatusByMsgId/<str:msgId>/', views.EditInErrorStatusByMsgId.as_view()),
# # -----------------------------------------------------------------------------------------------------------------#
#     # Excel Templatedownload
#     path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
#     #Excel Item Upload
    path('uploadedInItemExcel/', views.InItemExcelUpload.as_view()),
#     #Excel Item Upload
    path('editInAllItems/', views.InAllItemUpdate.as_view()),
#     #Delete Hawb by PermitId
    path('deleteInHawbByPermitId/<str:permit_id>/', views.DeleteInHawbl.as_view()),
#     # Inpaymentnew

#   # Excel Templatedownload
#     path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
#     #Excel Item Upload
#     path('uploadedExcelItem/', views.ItemExcelUpload.as_view()),
#     #Excel Item Upload
#     path('editAllItems/', views.AllItemUpdate.as_view()),
# # update item brand by permit id
    path('updateInItemBrand/', views.UpdateInItemBrandByPermitId.as_view()),
    #Delete Hawb by PermitId
    # path('deleteInHawbByPermitId/<str:permit_id>/', views.DeleteInHawbl.as_view()),
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
     path('gstInStatus/',views.GstInStatus.as_view()),
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

# Item Copy common

path('syncInItemFromCommon/', views.SyncInItemFromCommon.as_view()),

# copy refund




]

