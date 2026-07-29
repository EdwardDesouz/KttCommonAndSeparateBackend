from django.urls import path
from . import views,innonpayment,out,transhipment,coo

urlpatterns = [
    #login
    path('loginUser/', views.LoginAPIView.as_view()),
    path('logoutUser/', views.LogoutAPIView.as_view()),
    # path('currentUser/', views.CurrentUserAPIView.as_view()),
    #CommonHeaderTable
    path('getCommonHeaderTableInfo/', views.GetCommonHeaderTable.as_view()),
    path('postCommonHeaderTable/', views.PostCommonHeaderTable.as_view()),
    path("getCommonHeaderByPermitId/",views.GetCommonHeaderByPermitId.as_view()),
    path("editPermit/<str:permit_id>/", views.EditCommonHeaderByPermit.as_view()),
    path("deletePermit/", views.DeletePermit.as_view()),
    # Permit Details
    path("getDeclarantByMailbox/", views.GetDeclarantByMailbox.as_view()),
    # CopyPermit
    path("copyInpayment/", views.CopyInpayment.as_view()),
    #CommonInvoiceTable,
    path('getCommonInvoiceTableInfo/', views.CommonInvoiceTable.as_view()),
    path("getInvoiceNo/<str:invoice_no>/<str:permit_id>/",views.GetInvoiceByInvoiceNo.as_view()),
    path("getInvoiceByEditPermitId/",views.GetInvoiceByEditPermitId.as_view()),
    path("getInvoiceByPermitId/<str:permit_id>/",views.GetInvoiceByPermitId.as_view()),
    path("deleteInvoiceNo/",views.DeleteInvoice.as_view()), 
    path('postInvoiceTable/', views.PostInvoiceTable.as_view()),
    path('editInvoice/<str:invoice_no>/<str:permit_id>/', views.EditInvoiceByInvoiceNo.as_view()),
    #CommonItemTable
    path('getCommonItemTableInfo/', views.GetCommonItemTabel.as_view()),
    path('getItemNo/<str:item_no>/<str:permit_id>/', views.GetCommonItemByItemNo.as_view()),
    path("getItemByEditPermitId/",views.GetItemByEditPermitId.as_view()),
    path("deleteItem/",views.DeleteItem.as_view()),
    path('postItemTable/', views.PostItemTable.as_view()),
    path('editItem/<str:item_no>/<str:permit_id>/', views.EditItemByItemNo.as_view()),
    #casc
    path('getCommonCascTableInfo/', views.GetCommonCascTabel.as_view()),
    path('getCasc/<str:permit_id>/', views.GetCommonCascByPermitId.as_view()),
    path('deleteCasc/', views.DeleteCasc.as_view()),
    path('deleteCascByCascId/<str:casc_id>/<int:row_no>/<str:permit_id>/',views.DeleteCascByCascId.as_view()),
    path("postCascTable/", views.PostCascTable.as_view()),
    path('editCasc/<str:permit_id>/', views.EditCascByPermitId.as_view()),
    #cpc
    path('getCommonCpcTableInfo/', views.GetCommonCpcTabel.as_view()),
    path("getCpcByEditPermitId/",views.GetCommonCpcByEditPermitId.as_view()),
    path('getCpc/<str:permit_id>/', views.GetCommonCpcByPermitId.as_view()),
    path("deleteCpc/<str:permit_id>/", views.DeleteCpc.as_view()),
    path("postCpcTable/", views.PostCpcTable.as_view()),
    path('editCpc/<str:permit_id>/', views.EditCpcByPermitId.as_view()),
    #container
    path('getCommonContainerTableInfo/', views.GetCommonContainerTabel.as_view()),
    path("getContainerByEditPermitId/",views.GetContainerByEditPermitId.as_view()),
    path('getContainer/<str:permit_id>/', views.GetCommonContainerByPermitId.as_view()),
    path("deleteContainer/", views.DeleteContainer.as_view()),
    path("postContainerTable/", views.PostContainerTable.as_view()),
    path('editContainer/<str:permit_id>/', views.EditContainerByPermitId.as_view()),
    # Duplicate Hbl,Hawb
    path('checkDuplicateHawb/', views.CheckDuplicateHawb.as_view()),
    #housecode
    path('getCommonHouseItemCode/',views.GetCommonHouseItemCode.as_view()),
    path('getCommonHouseItemCodeByHouseCode/<str:housecode>/',views.GetCommonHouseItemCodeByHouseCode.as_view()),
    path('postCommonHouseItemCode/',views.PostCommonHouseItemCode.as_view()),
    path('editCommonHouseItemCode/<str:housecode>/',views.EditCommonHouseItemCode.as_view()),
    path('deleteCommonHouseItemCode/<str:housecode>/',views.DeleteCommonHouseItemCode.as_view()),
    #importer
    path('getCommonImporterTableInfo/', views.GetCommonImporterTabel.as_view()),
    path('getImporter/<str:code>/', views.GetCommonImporterByCode.as_view()),
    path("deleteImporter/<str:code>/", views.DeleteImporter.as_view()),
    path("postImporterTable/", views.PostImporterTable.as_view()),
    path('editImporter/<str:code>/', views.EditImporterByCode.as_view()),
    #HandlingAgent
    path('getCommonHandlingAgentTableInfo/', views.GetCommonHandlingAgentTabel.as_view()),
    path('getHandlingAgent/<str:code>/', views.GetCommonHandlingAgentByCode.as_view()),
    path("deleteHandlingAgent/<str:code>/", views.DeleteHandlingAgent.as_view()),
    path("postHandlingAgentTable/", views.PostHandlingAgentTable.as_view()),
    path('editHandlingAgent/<str:code>/', views.EditHandlingAgentByCode.as_view()),
    #FreightForwarder
    path("getCommonFreightForwarderTable/",views.GetCommonFreightForwarderTable.as_view()),
    path("getFreightForwarder/<str:code>/",views.GetFreightForwarderByCode.as_view()),
    path("deleteFreightForwarder/<str:code>/",views.DeleteFreightForwarder.as_view()),
    path("postFreightForwarderTable/",views.PostFreightForwarderTable.as_view()),
    path("editFreightForwarder/<str:code>/",views.EditFreightForwarderByCode.as_view()),
    #ClaimantParty
    path("getCommonClaimantPartyTable/", views.GetCommonClaimantPartyTable.as_view()),
    path("getClaimantParty/<int:id>/", views.GetClaimantPartyById.as_view()),
    path("deleteClaimantParty/<int:id>/", views.DeleteClaimantParty.as_view()),
    path("postClaimantPartyTable/", views.PostClaimantPartyTable.as_view()),
    path("editClaimantParty/<int:id>/", views.EditClaimantPartyById.as_view()), 
   
    #Exporter
    path('getCommonExporterTableInfo/', views.GetCommonExporterTabel.as_view()),
    path('getExporter/<str:code>/', views.GetCommonExporterByCode.as_view()),
    path("deleteExporter/<str:code>/", views.DeleteExporter.as_view()),
    path("postExporterTable/", views.PostExporterTable.as_view()),
    path('editExporter/<str:code>/', views.EditExporterByCode.as_view()),
    #Inward
    path('getCommonInwardCarrierAgentTableInfo/', views.CommonInwardCarrierAgent.as_view()),
    path('getInwardCarrierAgent/<str:code>/', views.GetCommonInwardCarrierAgentByCode.as_view()),
    path("deleteInwardCarrierAgent/<str:code>/", views.DeleteInwardCarrierAgent.as_view()),
    path("postInwardCarrierAgentTable/", views.PostInwardCarrierAgentTable.as_view()),
    path('editInward/<str:code>/', views.EditInwardCarrierAgentByCode.as_view()),
    #Outward
    path('getCommonOutwardCarrierAgentTableInfo/', views.CommonOutwardCarrierAgent.as_view()),
    path('getOutwardCarrierAgent/<str:code>/', views.GetCommonOutwardCarrierAgentByCode.as_view()),
    path("deleteOutwardCarrierAgent/<str:code>/", views.DeleteOutwardCarrierAgent.as_view()),
    path("postOutwardCarrierAgentTable/", views.PostOutwardCarrierAgentTable.as_view()),
    path('editOutward/<str:code>/', views.EditOutwardCarrierAgentByCode.as_view()),
    #Consignee
    path('getCommonConsigneeTableInfo/', views.CommonConsigneeTable.as_view()),
    path('getConsignee/<str:consigneecode>/', views.GetCommonConsigneeAgentByCode.as_view()),
    path("deleteCongineeAgent/<str:consigneecode>/", views.DeleteConsigneeAgentByCode.as_view()),
    path("postCongineeTable/", views.PostCongineeTable.as_view()),
    path('editConginee/<str:consigneecode>/', views.EditCongineeCode.as_view()),
    #EndUser
    path('getCommonEndUserTableInfo/', views.CommonEndUserTable.as_view()),
    path('getEndUser/<str:EndUserCode>/', views.GetCommonEndUserByCode.as_view()),
    path("deleteEndUser/<str:EndUserCode>/", views.DeleteEndUserByCode.as_view()),
    path("postEndUserTable/", views.PostEndUserTable.as_view()),
    path('editEndUser/<str:EndUserCode>/', views.EditEndUserCode.as_view()),
    #Manufactrer
    path('getCommonManufacturerTableInfo/',views.CommonManufacturerTable.as_view()),
    path('getCommonManufacturerByCode/<str:ManufacturerCode>/',views.GetCommonManufacturerByCode.as_view()),
    path('postManufacturerTable/',views.PostManufacturerTable.as_view()),
    path('editManufacturerCode/<str:ManufacturerCode>/',views.EditManufacturerCode.as_view()),
    path('deleteManufacturerByCode/<str:ManufacturerCode>/',views. DeleteManufacturerByCode.as_view()),
    #CommonSupplierManufacturerPart
    path('getCommonSupplierManufacturerPartTableInfo/', views.GetCommonSupplierManufacturerPart.as_view()),
    path('getSupplierManufacturerPart/<str:code>/', views.GetCommonSupplierManufacturerPartByCode.as_view()),
    path('deleteSupplierManufacturerPart/<str:code>/', views.DeleteCommonSupplierManufacturerPart.as_view()),
    path('postSupplierManufacturerPartTable/', views.PostCommonSupplierManufacturerPart.as_view()),
    path('editSupplierManufacturerPart/<str:code>/', views.EditCommonSupplierManufacturerPartByCode.as_view()),

    #file
    path('getCommonFileTableInfo/', views.GetFileTabel.as_view()),
    path("getCommonFileByEditPermitId/",views.GetCommonFileByEditPermitId.as_view()),
    path('getFile/<str:permit_id>/', views.GetCommonFileByPermitId.as_view()),
    path("deleteFile/<str:permit_id>/<int:sno>/",views.DeleteFile.as_view()),
    path("postFileTable/", views.PostFileTable.as_view()),
    path('editFile/<str:permit_id>/', views.EditFileByPermitId.as_view()),
    #PMT
    path('getCommonPmtTableInfo/', views.GetPMTTabel.as_view()),
    path('getPmt/<str:permit_number>/', views.GetCommonPMTByPermitNo.as_view()),
    path("deletePmt/<str:permit_number>/", views.DeletePmt.as_view()),
    path("postPmtTable/", views.PostPmtTable.as_view()),
    path('editPmt/<str:permit_number>/', views.EditPmtByPermitNo.as_view()),
   #AMdPMT
    path('getCommonAmdPmtTableInfo/', views.GetAMDPMTTabel.as_view()),
    path('getAmdPmt/<str:permit_number>/', views.GetCommonAMDPMTByPermitNo.as_view()),
    path("deleteAmdPmt/<str:permit_number>/", views.DeleteAMDPmt.as_view()),
    path("postAmdPmtTable/", views.PostAMDPmtTable.as_view()),
    path('editAmdPmt/<str:permit_number>/', views.EditAMDPmtByPermitNo.as_view()),
   #RejectStatus
    path('getCommonRejectTableInfo/', views.CommonRejectStausTable.as_view()),
    path('getRejectStatusMessage/<str:msgId>/', views.GetCommonRejectStautsByMsgId.as_view()),
    path("deleteRejectStatus/<str:msgId>/", views.DeleteRejectStatusByMsgId.as_view()),
    path("postRejectStatusTable/", views.PostRejectStautsTable.as_view()),
    path('editMessage/<str:msgId>/', views.EditRejectStatusByMsgId.as_view()),
    #ErrorStatus
    path('getCommonErrorTableInfo/', views.CommonErrorStatusTable.as_view()),
    path('getRejectErrorMessage/<str:msgId>/', views.GetCommonErrorStautsByMsgId.as_view()),
    path("deleteErrorStatus/<str:msgId>/", views.DeleteErrorStatusByMsgId.as_view()),
    path("postErrorStatusTable/", views.PostErrorStautsTable.as_view()),
    path('editerror/<str:msgId>/', views.EditErrorStatusByMsgId.as_view()),
    #CommonDropDownTable
    #CommonMaster
    path('getCommonMasterTableInfo/', views.GetCommonMasterTable.as_view()),
    #Currency
    path('getCommonCurrencyTableInfo/', views.GetCommonCurrenyTable.as_view()),
    #Country
    path('getCommonCountryTableInfo/', views.GetCommonCountryTable.as_view()),
    #DeclarantCompany
    path('getCommonDeclarantCompanyTableInfo/', views.GetCommonDeclarantCompanyTable.as_view()),
    #Hscode
    path('getCommonHsCodeTableInfo/', views.GetCommonHsCodeTable.as_view()),
    #Hscode
    path('getCommonHsCodeAndDescriptionTable/', views.GetCommonHsCodeAndDescriptionTable.as_view()),
    #UOM
    path('getUOMFromCommonHscode/', views.GetUOMFromCommonHscode.as_view()),
    #TermType
    path('getTermTypeFromCommonMaster/', views.GetTermTypeFromCommonMaster.as_view()),
    #TotalOuterPack
    path('getTotalOuterPackFromCommonMaster/', views.GetTotalOuterPackFromCommonMaster.as_view()),
    #DECLARINGFOR
    path('getDeclaringForFromCommonMaster/', views.GetDeclaringForFromCommonMaster.as_view()),
    #DECLARINGFOR
    path('getDeclaringForFromCommonMasterByOutandTranshipment/', views.GetDeclaringForFromCommonMasterByOutandTranshipment.as_view()),
    #InwardTransportMode
    path('getInwardTransportModeFromCommonMaster/', views.GetInwardTransportModeFromCommonMaster.as_view()),
    #DeclarationType
    path('getDeclarationTypeFromCommonMaster/', views.GetDeclarationTypeFromCommonMaster.as_view()),
    #DeclarationType--inpaymetn
    path('getDeclarationTypeFromCommonMasterForInpayment/', views.GetDeclarationTypeFromCommonMasterForInpayment.as_view()),
    #DeclarationType--innonpayment
    path('getDeclarationTypeFromCommonMasterForInnonpayment/', views.GetDeclarationTypeFromCommonMasterForInnonpayment.as_view()),
    #DeclarationType--out
    path('getDeclarationTypeFromCommonMasterForOut/', views.GetDeclarationTypeFromCommonMasterForOut.as_view()),
    #DeclarationType--transhipment
    path('getDeclarationTypeFromCommonMasterForTranshipment/', views.GetDeclarationTypeFromCommonMasterForTranshipment.as_view()),
    #Co-Type --Out
    path('getCoTypeFromCommonMasterForOut/', views.GetCoTypeFromCommonMasterForOut.as_view()),
    #BgIndicator
    path('getBgIndicatorFromCommonMaster/', views.GetBgIndicatorFromCommonMaster.as_view()),
    #Preferntial
    path('getPreferntialFromCommonMaster/', views.GetPreferntialFromCommonMaster.as_view()),
    #VehicalType
    path('getVehicalTypeFromCommonMaster/', views.GetVehicalTypeFromCommonMaster.as_view()),
    #EngineCapacity
    path('getEngineCapacityFromCommonMaster/', views.GetEngineCapacityFromCommonMaster.as_view()),
    #MakingLot
    path('getMakingLotFromCommonMaster/', views.GetMakingLotFromCommonMaster.as_view()),
    #CancelType
    path('getCancelTypeFromCommonMaster/', views.GetCancelTypeFromCommonMaster.as_view()),
    #RefundType
    path('getRefundTypeFromCommonMaster/', views.GetRefundTypeFromCommonMaster.as_view()),
    #ReasonForRefund
    path('getReasonForRefundFromCommonMaster/', views.GetReasonForRefundFromCommonMaster.as_view()),
    #CoType
    path('getCoTypeFromCommonMaster/', views.GetCoTypeFromCommonMaster.as_view()),
    #DocumentAttachmentType
    path('getDocumentAttachFromCommonMaster/', views.GetDocumentAttachTypeFromCommonMaster.as_view()),
    #CertificateType
    path('getCertificateTypeFromCommonMaster/', views.GetCertificateTypeFromCommonMaster.as_view()),
    #Container
    path('getContainerFromCommonMaster/', views.GetContainerFromCommonMaster.as_view()),
    #Making
    path('getMakingFromCommonMaster/', views.GetMakingFromCommonMaster.as_view()),
    #VesselType
    path('getVesselTypeFromCommonMaster/', views.GetVesselTypeFromCommonMaster.as_view()),
    #Making
    path('getCargoTypeFromCommonMaster/', views.GetCargoTypeFromCommonMaster.as_view()),
    #EngineCapacity
    path('getEngineCapacityFromCommonMaster/', views.GetEngineCapacityFromCommonMaster.as_view()),
    #ReleaseLocation
    path('getReleaseLocation/', views.GetReleaseLocation.as_view()),
    #ReceiptLocation
    path('getReceiptLocation/', views.GetReceiptLocation.as_view()),
    #LoadingPort
    path('getLoadingPort/', views.GetLoadingPort.as_view()),
    #StorageLocation
    path('getStorageLocation/', views.GetStorageLocation.as_view()),
    #ManageUser
    path('getManageUserMail/', views.GetManageUserMail.as_view()),
    # Product Code By HSCode
    path("getCascProductCodes/", views.GetCascProductCodes.as_view()),
# -----------------------------------------------------------------------------------------------------------------#
    # Excel Templatedownload
    path('downloadExcelTemplate/<str:template_name>/', views.ItemExcelDownload.as_view()),
    #Excel Item Upload
    path('uploadedExcelItem/', views.ItemExcelUpload.as_view()),
    #Excel Item Upload
    path('editAllItems/', views.AllItemUpdate.as_view()),
# update item brand by permit id
   path('updateItemBrand/', views.UpdateItemBrandByPermitId.as_view()),
    #Delete Hawb by PermitId
    path('deleteHawbByPermitId/<str:permit_id>/', views.DeleteHawbl.as_view()),
    # Inpaymentnew
    path('inpaymentnew/',views.InpaymentNewPermit.as_view()),
    path('inpaymentList/',views.InpaymentList.as_view()),
    # PrintGst
    path('PrintGst/<str:PermitId>/', views.PrintGst.as_view()),
    # PrintGst
    path('printGstAll/', views.PrintGstAll.as_view()),
    # Download ccp
    path('downloadCcp/', views.DownloadCcp.as_view()),
    # PRINT CCP
    path('printCcp/<str:permit_id>/',views.PrintCcp.as_view()),
    # Gst status
    path('gstStatus/',views.GstStatus.as_view()),
    # Download Data
    path('downloadData/', views.DownloadData.as_view()),
    # Print Gst Via Excel
    path('gstExcel/', views.GstExcel.as_view()),
    # Print Status
    path('printStatus/<str:PermitId>/', views.PrintStatus.as_view()),
    # Xml Submit 
    path('XmlSubmit/', views.XmlSubmit.as_view()),
    # Amend
    path('getAmendPermitByMsgId/', views.GetAmendByMsgId.as_view()),
    path('postAmendTable/', views.PostAmendTable.as_view()),
    # Cancel
    path('getCancelPermitByMsgId/', views.CancelPermit.as_view()),
    path('postCancelPermit/', views.PostCancelPermit.as_view()),
    # Refund
    path('getRefundPermitByMsgId/', views.RefundPermit.as_view()),
    path('postRefundPermit/', views.PostRefundPermit.as_view()),
    path('getRefundValSummaryByMsgId/', views.RefundValSummary.as_view()),
    path('postRefundValSummary/', views.PostRefundValSummary.as_view()),
    path('getRefundItemSummaryByMsgId/', views.GetReundItemSummByMsgId.as_view()),
    path('postRefundItemSummary/', views.PostReundItemSumm.as_view()),
    # get permit conditions for view 
    path('getPermitConditions/', views.GetPermitConditions.as_view()),
    # transmit innonpayment
    path("transmitInnonpayment/", views.TransmitInnonpayment.as_view()),
    # mailbox transmit
    path("mailboxTransmitData/", views.MailBoxTransmitData.as_view()),
    # Get Exchange Rate using Date
    path("getExchangeRateByDate/",views.GetExchangeRateByDate.as_view()),


    #Innonpayment
    # List page
    path('innonpaymentList/',innonpayment.InnonpaymentList.as_view()),
    # New Permit
    path('innonpaymentNew/',innonpayment.InnonpaymentNewPermit.as_view()),
    # Copy Permit
    path("copyInnonpayment/", innonpayment.CopyInnonpayment.as_view()),
    # transmit innonpayment
    path("transmitInpayment/", innonpayment.TransmitInpayment.as_view()),

    # Out
    # List Page
    path('outList/',out.OutList.as_view()),
    # New Permit
    path('outNew/',out.OutNewPermit.as_view()),
    # Copy Permit
    path("copyOut/", out.CopyOut.as_view()),
    # Draft Coo
    path("draftCoo/<str:PermitId>/", out.DraftCoo.as_view()),
    # Print Coo
    path("printCoo/<str:PermitId>/", out.PrintCoo.as_view()),
     # transmit innonpayment
    path("transmitOut/", out.TransmitOutpayment.as_view()),


    # Transhipment
    # List Page
    path('transList/',transhipment.TransList.as_view()),
    # New Permit
    path('transNew/',transhipment.TranshipmentNewPermit.as_view()),
    # Copy Permit
    path("copyTranshipment/", transhipment.CopyTranshipment.as_view()),
    
    # Coo
    # List Page
    path("cooList/",coo.CooList.as_view()),
    # New Permit
    path('cooNew/',coo.CooNewPermit.as_view()),
    # Copy Coo
    path("copyCoo/", coo.CopyCoo.as_view()),

    # Inpayment
    # Refund
    path("copyInpaymentRefund/",views.CopyRefund.as_view()),
    # Amend
    path('copyInpaymentAmend/',views.CopyAmend.as_view()),
    # cancel
    path('copyInpaymentCancel/',views.CopyCancel.as_view()),

    # Innonpayment
    # Amend
    path('copyInnonAmend/',innonpayment.CopyInnonAmend.as_view()),
    # POSTINNONDAMMEND
    path('postInnonAmendTable/',innonpayment.PostInnonAmendTable.as_view()),
    # Cancel
    path('copyInnonpaymentCancel/',innonpayment.CopyInnonCancel.as_view()),
    # PostInnonCancel
    path('postInnonCancelTable/',innonpayment.PostInnonCancelTable.as_view()),


    # Out
    # Amend
    path('copyOutAmend/',out.CopyOutAmend.as_view()),
    # POSTOutAMMEND
    path('postOutAmendTable/',out.PostOutAmendTable.as_view()),
    # Cancel
    path('copyOutCancel/',out.CopyOutCancel.as_view()),
    # PostInnonCancel
    path('postOutCancelTable/',out.PostOutCancelTable.as_view()),

    # Transhipment
    # Amend
    path('copyTransAmend/',transhipment.CopyTransAmend.as_view()),
    # POSTOutAMMEND
    path('postTransAmendTable/',transhipment.PostTransAmendTable.as_view()),

 # COPY CANCEL
    path('copyTransCancel/',transhipment.CopyCancelTranshipment.as_view()),
    # POSTCancel
    path('postTransCancelTable/',transhipment.PostCancelTranshipment.as_view()),

   


]
