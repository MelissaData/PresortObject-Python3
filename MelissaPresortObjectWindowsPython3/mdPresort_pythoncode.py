from ctypes import *
from enum import Enum
import ctypes
import os
import sys

if (os.name == 'nt' and sys.version_info[:2] >= (3,8)):
  lib = ctypes.CDLL('mdPresort.dll', winmode=0)
elif (os.name == 'nt'):
  lib = ctypes.CDLL('mdPresort.dll')
else:
  lib = ctypes.CDLL('libmdPresort.so')

lib.mdPresortCreate.argtypes = []
lib.mdPresortCreate.restype = c_void_p
lib.mdPresortDestroy.argtypes = [c_void_p]
lib.mdPresortDestroy.restype = None
lib.mdPresortInitializeDataFiles.argtypes = [c_void_p]
lib.mdPresortInitializeDataFiles.restype = c_int
lib.mdPresortSetLicenseString.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetLicenseString.restype = c_bool
lib.mdPresortSetPathToPresortDataFiles.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPathToPresortDataFiles.restype = None
lib.mdPresortGetInitializeErrorString.argtypes = [c_void_p]
lib.mdPresortGetInitializeErrorString.restype = c_char_p
lib.mdPresortGetParametersErrorString.argtypes = [c_void_p]
lib.mdPresortGetParametersErrorString.restype = c_char_p
lib.mdPresortGetBuildNumber.argtypes = [c_void_p]
lib.mdPresortGetBuildNumber.restype = c_char_p
lib.mdPresortGetLicenseStringExpirationDate.argtypes = [c_void_p]
lib.mdPresortGetLicenseStringExpirationDate.restype = c_char_p
lib.mdPresortGetDatabaseExpirationDate.argtypes = [c_void_p]
lib.mdPresortGetDatabaseExpirationDate.restype = c_char_p
lib.mdPresortGetDatabaseDate.argtypes = [c_void_p]
lib.mdPresortGetDatabaseDate.restype = c_char_p
lib.mdPresortSetZip.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetZip.restype = None
lib.mdPresortSetPlus4.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPlus4.restype = None
lib.mdPresortSetCarrierRoute.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetCarrierRoute.restype = None
lib.mdPresortSetWalkSequence.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetWalkSequence.restype = None
lib.mdPresortSetDeliveryPointCode.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetDeliveryPointCode.restype = None
lib.mdPresortSetBusinessResidentialIndicator.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetBusinessResidentialIndicator.restype = None
lib.mdPresortSetPSPermitImprint.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSPermitImprint.restype = None
lib.mdPresortSetLOTNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetLOTNumber.restype = None
lib.mdPresortSetLOTOrder.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetLOTOrder.restype = None
lib.mdPresortSetPSASE.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSASE.restype = None
lib.mdPresortSetPSFASTforward.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSFASTforward.restype = None
lib.mdPresortSetPSNCOA.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSNCOA.restype = None
lib.mdPresortSetPSACS.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSACS.restype = None
lib.mdPresortSetPSAltMethod.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSAltMethod.restype = None
lib.mdPresortSetPSMultiple.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSMultiple.restype = None
lib.mdPresortSetPSOneCode.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSOneCode.restype = None
lib.mdPresortSetPSAltAddFmt.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSAltAddFmt.restype = None
lib.mdPresortSetPSCASSDate.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSCASSDate.restype = None
lib.mdPresortSetPSPrecanceledStamp.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSPrecanceledStamp.restype = None
lib.mdPresortSetPSPermitHolderName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderName.restype = None
lib.mdPresortSetPSPermitHolderCompany.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderCompany.restype = None
lib.mdPresortSetPSPermitHolderStreet.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderStreet.restype = None
lib.mdPresortSetPSPermitHolderCity.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderCity.restype = None
lib.mdPresortSetPSPermitHolderState.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderState.restype = None
lib.mdPresortSetPSPermitHolderPhone.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderPhone.restype = None
lib.mdPresortSetPSPermitHolderEmail.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderEmail.restype = None
lib.mdPresortSetPSPermitHolderListName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderListName.restype = None
lib.mdPresortSetPSPermitHolderZIP.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPermitHolderZIP.restype = None
lib.mdPresortSetPSPostOfficeOfMailingCity.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPostOfficeOfMailingCity.restype = None
lib.mdPresortSetPSPostOfficeOfMailingState.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPostOfficeOfMailingState.restype = None
lib.mdPresortSetPSPostOfficeOfMailingZIP.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPostOfficeOfMailingZIP.restype = None
lib.mdPresortSetPSMailingAgentName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentName.restype = None
lib.mdPresortSetPSMailingAgentCompany.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentCompany.restype = None
lib.mdPresortSetPSMailingAgentStreet.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentStreet.restype = None
lib.mdPresortSetPSMailingAgentCity.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentCity.restype = None
lib.mdPresortSetPSMailingAgentState.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentState.restype = None
lib.mdPresortSetPSMailingAgentPhone.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentPhone.restype = None
lib.mdPresortSetPSMailingAgentZIP.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentZIP.restype = None
lib.mdPresortSetPSIndividualOrOrganizationName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSIndividualOrOrganizationName.restype = None
lib.mdPresortSetPSIndividualOrOrganizationCompany.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSIndividualOrOrganizationCompany.restype = None
lib.mdPresortSetPSIndividualOrOrganizationStreet.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSIndividualOrOrganizationStreet.restype = None
lib.mdPresortSetPSIndividualOrOrganizationCity.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSIndividualOrOrganizationCity.restype = None
lib.mdPresortSetPSIndividualOrOrganizationState.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSIndividualOrOrganizationState.restype = None
lib.mdPresortSetPSIndividualOrOrganizationZIP.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSIndividualOrOrganizationZIP.restype = None
lib.mdPresortSetPSIndividualOrOrganizationCRID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSIndividualOrOrganizationCRID.restype = None
lib.mdPresortSetContinueContainerNumber.argtypes = [c_void_p, c_bool]
lib.mdPresortSetContinueContainerNumber.restype = None
lib.mdPresortSetPOMasNDC.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPOMasNDC.restype = c_bool
lib.mdPresortSetPOMasSCF.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPOMasSCF.restype = c_bool
lib.mdPresortSetPSStatementSeqNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSStatementSeqNumber.restype = None
lib.mdPresortSetPSMailingDate.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingDate.restype = None
lib.mdPresortSetPSFedAgencyCode.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSFedAgencyCode.restype = None
lib.mdPresortSetPSCustomerNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSCustomerNumber.restype = None
lib.mdPresortSetPSCAPSNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSCAPSNumber.restype = None
lib.mdPresortSetPermitNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPermitNumber.restype = None
lib.mdPresortSetPSNonProfitAuthNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSNonProfitAuthNumber.restype = None
lib.mdPresortSetMailClass.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMailClass.restype = None
lib.mdPresortSetPieceType.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPieceType.restype = None
lib.mdPresortSetSortType.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetSortType.restype = None
lib.mdPresortSetFCM_CRRT.argtypes = [c_void_p, c_bool]
lib.mdPresortSetFCM_CRRT.restype = None
lib.mdPresortSetFCM_CRRT_5.argtypes = [c_void_p, c_bool]
lib.mdPresortSetFCM_CRRT_5.restype = None
lib.mdPresortSetFCM_CRRT_3.argtypes = [c_void_p, c_bool]
lib.mdPresortSetFCM_CRRT_3.restype = None
lib.mdPresortSetFCM_5dg_Scheme.argtypes = [c_void_p, c_bool]
lib.mdPresortSetFCM_5dg_Scheme.restype = None
lib.mdPresortSetFCM_Auto_5dg.argtypes = [c_void_p, c_bool]
lib.mdPresortSetFCM_Auto_5dg.restype = None
lib.mdPresortSetFCM_NonAuto_5dg.argtypes = [c_void_p, c_bool]
lib.mdPresortSetFCM_NonAuto_5dg.restype = None
lib.mdPresortSetSTD_Auto_5dg_Scheme.argtypes = [c_void_p, c_bool]
lib.mdPresortSetSTD_Auto_5dg_Scheme.restype = None
lib.mdPresortSetSTD_Auto_5dg.argtypes = [c_void_p, c_bool]
lib.mdPresortSetSTD_Auto_5dg.restype = None
lib.mdPresortSetNonMachinableType.argtypes = [c_void_p, c_bool]
lib.mdPresortSetNonMachinableType.restype = None
lib.mdPresortSetRecordID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetRecordID.restype = None
lib.mdPresortSetSackWeight.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetSackWeight.restype = None
lib.mdPresortSetPieceLength.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPieceLength.restype = None
lib.mdPresortSetPieceHeight.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPieceHeight.restype = None
lib.mdPresortSetPieceThickness.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPieceThickness.restype = None
lib.mdPresortSetPieceWeight.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPieceWeight.restype = None
lib.mdPresortSetPackageSize.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPackageSize.restype = None
lib.mdPresortSetEnableCoTray.argtypes = [c_void_p, c_bool]
lib.mdPresortSetEnableCoTray.restype = None
lib.mdPresortSetUseFSM1000.argtypes = [c_void_p, c_bool]
lib.mdPresortSetUseFSM1000.restype = None
lib.mdPresortSetSTDNonProfit.argtypes = [c_void_p, c_bool]
lib.mdPresortSetSTDNonProfit.restype = None
lib.mdPresortSetIgnoreDSF.argtypes = [c_void_p, c_bool]
lib.mdPresortSetIgnoreDSF.restype = None
lib.mdPresortSetMailersID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMailersID.restype = None
lib.mdPresortSetPSPrecanceledStampValue.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPrecanceledStampValue.restype = c_bool
lib.mdPresortSetPSPresortResidualPieces.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSPresortResidualPieces.restype = None
lib.mdPresortSetIMBSerialNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetIMBSerialNumber.restype = None
lib.mdPresortSetPSPolicitalCampaignMailing.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSPolicitalCampaignMailing.restype = None
lib.mdPresortSetPSOfficialElectionMail.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSOfficialElectionMail.restype = None
lib.mdPresortProduceReports.argtypes = [c_void_p, c_char_p, c_char_p]
lib.mdPresortProduceReports.restype = c_bool
lib.mdPresortSetSCFCity.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetSCFCity.restype = None
lib.mdPresortSetSCFState.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetSCFState.restype = None
lib.mdPresortSetSCFZip.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetSCFZip.restype = None
lib.mdPresortAddSCF.argtypes = [c_void_p]
lib.mdPresortAddSCF.restype = c_bool
lib.mdPresortSetNDCCity.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetNDCCity.restype = None
lib.mdPresortSetNDCState.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetNDCState.restype = None
lib.mdPresortSetNDCZip.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetNDCZip.restype = None
lib.mdPresortAddNDC.argtypes = [c_void_p]
lib.mdPresortAddNDC.restype = c_bool
lib.mdPresortSetDDUCity.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetDDUCity.restype = None
lib.mdPresortSetDDUState.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetDDUState.restype = None
lib.mdPresortSetDDUZip.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetDDUZip.restype = None
lib.mdPresortSetDDUMoreZip.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetDDUMoreZip.restype = None
lib.mdPresortAddDDU.argtypes = [c_void_p]
lib.mdPresortAddDDU.restype = c_bool
lib.mdPresortSetTTBorder.argtypes = [c_void_p, c_bool]
lib.mdPresortSetTTBorder.restype = None
lib.mdPresortSetTTNumberofPieces.argtypes = [c_void_p, c_bool]
lib.mdPresortSetTTNumberofPieces.restype = None
lib.mdPresortSetTTContainerNumber.argtypes = [c_void_p, c_bool]
lib.mdPresortSetTTContainerNumber.restype = None
lib.mdPresortSetTTContainerSize.argtypes = [c_void_p, c_bool]
lib.mdPresortSetTTContainerSize.restype = None
lib.mdPresortSetTTOther.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetTTOther.restype = None
lib.mdPresortSetTTParameterPositionX.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetTTParameterPositionX.restype = None
lib.mdPresortSetTTParameterPositionY.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetTTParameterPositionY.restype = None
lib.mdPresortSetTTParameterWidth.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetTTParameterWidth.restype = None
lib.mdPresortSetTTParameterHeight.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetTTParameterHeight.restype = None
lib.mdPresortDoPresort.argtypes = [c_void_p]
lib.mdPresortDoPresort.restype = c_bool
lib.mdPresortSetDisableDetailedOutput.argtypes = [c_void_p, c_bool]
lib.mdPresortSetDisableDetailedOutput.restype = None
lib.mdPresortSetPreSortSettings.argtypes = [c_void_p, c_long]
lib.mdPresortSetPreSortSettings.restype = c_char_p
lib.mdPresortSetACSCodeSettings.argtypes = [c_void_p, c_long]
lib.mdPresortSetACSCodeSettings.restype = c_bool
lib.mdPresortUpdateParameters.argtypes = [c_void_p]
lib.mdPresortUpdateParameters.restype = c_bool
lib.mdPresortAddRecord.argtypes = [c_void_p]
lib.mdPresortAddRecord.restype = c_bool
lib.mdPresortGetRecord.argtypes = [c_void_p, c_char_p]
lib.mdPresortGetRecord.restype = c_bool
lib.mdPresortGetFirstRecord.argtypes = [c_void_p]
lib.mdPresortGetFirstRecord.restype = c_bool
lib.mdPresortGetNextRecord.argtypes = [c_void_p]
lib.mdPresortGetNextRecord.restype = c_bool
lib.mdPresortGetMailPieceRate.argtypes = [c_void_p]
lib.mdPresortGetMailPieceRate.restype = c_char_p
lib.mdPresortGetZipAsString.argtypes = [c_void_p]
lib.mdPresortGetZipAsString.restype = c_char_p
lib.mdPresortGetTrayNumber.argtypes = [c_void_p]
lib.mdPresortGetTrayNumber.restype = c_long
lib.mdPresortGetSequenceNumber.argtypes = [c_void_p]
lib.mdPresortGetSequenceNumber.restype = c_long
lib.mdPresortGetRecordID.argtypes = [c_void_p]
lib.mdPresortGetRecordID.restype = c_char_p
lib.mdPresortGetEndorsementLine.argtypes = [c_void_p]
lib.mdPresortGetEndorsementLine.restype = c_char_p
lib.mdPresortGetRateCode.argtypes = [c_void_p]
lib.mdPresortGetRateCode.restype = c_char_p
lib.mdPresortGetMailJob.argtypes = [c_void_p]
lib.mdPresortGetMailJob.restype = c_char_p
lib.mdPresortGetSortLevel.argtypes = [c_void_p]
lib.mdPresortGetSortLevel.restype = c_char_p
lib.mdPresortGetLabelLine1.argtypes = [c_void_p]
lib.mdPresortGetLabelLine1.restype = c_char_p
lib.mdPresortGetLabelLine2.argtypes = [c_void_p]
lib.mdPresortGetLabelLine2.restype = c_char_p
lib.mdPresortGetLabelList.argtypes = [c_void_p]
lib.mdPresortGetLabelList.restype = c_char_p
lib.mdPresortGetCINCode.argtypes = [c_void_p]
lib.mdPresortGetCINCode.restype = c_char_p
lib.mdPresortGetTrayProcessingCode.argtypes = [c_void_p]
lib.mdPresortGetTrayProcessingCode.restype = c_char_p
lib.mdPresortGetMailSplitCode.argtypes = [c_void_p]
lib.mdPresortGetMailSplitCode.restype = c_char_p
lib.mdPresortGetTrayZipCode.argtypes = [c_void_p]
lib.mdPresortGetTrayZipCode.restype = c_char_p
lib.mdPresortGetTrayType.argtypes = [c_void_p]
lib.mdPresortGetTrayType.restype = c_char_p
lib.mdPresortGetBundleZipCode.argtypes = [c_void_p]
lib.mdPresortGetBundleZipCode.restype = c_char_p
lib.mdPresortGetBundleSortLevel.argtypes = [c_void_p]
lib.mdPresortGetBundleSortLevel.restype = c_char_p
lib.mdPresortGetBundleNumber.argtypes = [c_void_p]
lib.mdPresortGetBundleNumber.restype = c_long
lib.mdPresortGetBundleIndicator.argtypes = [c_void_p]
lib.mdPresortGetBundleIndicator.restype = c_char_p
lib.mdPresortGetZipCodeInScheme.argtypes = [c_void_p]
lib.mdPresortGetZipCodeInScheme.restype = c_char_p
lib.mdPresortGetBarcodeID.argtypes = [c_void_p]
lib.mdPresortGetBarcodeID.restype = c_char_p
lib.mdPresortGetServiceTypeID.argtypes = [c_void_p]
lib.mdPresortGetServiceTypeID.restype = c_char_p
lib.mdPresortSetProducePallets.argtypes = [c_void_p, c_bool]
lib.mdPresortSetProducePallets.restype = c_bool
lib.mdPresortSetIgnoreMinSCFPallet.argtypes = [c_void_p, c_bool]
lib.mdPresortSetIgnoreMinSCFPallet.restype = None
lib.mdPresortGetPalletsTotal.argtypes = [c_void_p]
lib.mdPresortGetPalletsTotal.restype = c_long
lib.mdPresortGetPalletZipCode.argtypes = [c_void_p]
lib.mdPresortGetPalletZipCode.restype = c_char_p
lib.mdPresortGetPalletLabelLine1.argtypes = [c_void_p]
lib.mdPresortGetPalletLabelLine1.restype = c_char_p
lib.mdPresortGetPalletLabelLine2.argtypes = [c_void_p]
lib.mdPresortGetPalletLabelLine2.restype = c_char_p
lib.mdPresortGetPalletSortLevel.argtypes = [c_void_p]
lib.mdPresortGetPalletSortLevel.restype = c_char_p
lib.mdPresortGetPalletNumber.argtypes = [c_void_p]
lib.mdPresortGetPalletNumber.restype = c_long
lib.mdPresortGetTotalPiecesAuto5dig.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesAuto5dig.restype = c_long
lib.mdPresortGetTotalPiecesAuto3dig.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesAuto3dig.restype = c_long
lib.mdPresortGetTotalPiecesAutoADC.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesAutoADC.restype = c_long
lib.mdPresortGetTotalPiecesAutoMADC.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesAutoMADC.restype = c_long
lib.mdPresortGetTotalPieces5dig.argtypes = [c_void_p]
lib.mdPresortGetTotalPieces5dig.restype = c_long
lib.mdPresortGetTotalPieces3dig.argtypes = [c_void_p]
lib.mdPresortGetTotalPieces3dig.restype = c_long
lib.mdPresortGetTotalPiecesADC.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesADC.restype = c_long
lib.mdPresortGetTotalPiecesMADC.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesMADC.restype = c_long
lib.mdPresortGetTotalPiecesWS.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesWS.restype = c_long
lib.mdPresortGetTotalPiecesHDP.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesHDP.restype = c_long
lib.mdPresortGetTotalPiecesHD.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesHD.restype = c_long
lib.mdPresortGetTotalPiecesLOT.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesLOT.restype = c_long
lib.mdPresortGetRateAuto5dig.argtypes = [c_void_p]
lib.mdPresortGetRateAuto5dig.restype = c_double
lib.mdPresortGetRateAuto3dig.argtypes = [c_void_p]
lib.mdPresortGetRateAuto3dig.restype = c_double
lib.mdPresortGetRateAutoADC.argtypes = [c_void_p]
lib.mdPresortGetRateAutoADC.restype = c_double
lib.mdPresortGetRateAutoMADC.argtypes = [c_void_p]
lib.mdPresortGetRateAutoMADC.restype = c_double
lib.mdPresortGetRate5dig.argtypes = [c_void_p]
lib.mdPresortGetRate5dig.restype = c_double
lib.mdPresortGetRate3dig.argtypes = [c_void_p]
lib.mdPresortGetRate3dig.restype = c_double
lib.mdPresortGetRateADC.argtypes = [c_void_p]
lib.mdPresortGetRateADC.restype = c_double
lib.mdPresortGetRateMADC.argtypes = [c_void_p]
lib.mdPresortGetRateMADC.restype = c_double
lib.mdPresortGetRateWS.argtypes = [c_void_p]
lib.mdPresortGetRateWS.restype = c_double
lib.mdPresortGetRateHDP.argtypes = [c_void_p]
lib.mdPresortGetRateHDP.restype = c_double
lib.mdPresortGetRateHD.argtypes = [c_void_p]
lib.mdPresortGetRateHD.restype = c_double
lib.mdPresortGetRateLOT.argtypes = [c_void_p]
lib.mdPresortGetRateLOT.restype = c_double
lib.mdPresortGetPoundRate.argtypes = [c_void_p]
lib.mdPresortGetPoundRate.restype = c_double
lib.mdPresortGetPoundRateWS.argtypes = [c_void_p]
lib.mdPresortGetPoundRateWS.restype = c_double
lib.mdPresortGetPoundRateHDP.argtypes = [c_void_p]
lib.mdPresortGetPoundRateHDP.restype = c_double
lib.mdPresortGetPoundRateHD.argtypes = [c_void_p]
lib.mdPresortGetPoundRateHD.restype = c_double
lib.mdPresortGetPoundRateLOT.argtypes = [c_void_p]
lib.mdPresortGetPoundRateLOT.restype = c_double
lib.mdPresortGetTotalTrays1Ft.argtypes = [c_void_p]
lib.mdPresortGetTotalTrays1Ft.restype = c_long
lib.mdPresortGetTotalTrays2Ft.argtypes = [c_void_p]
lib.mdPresortGetTotalTrays2Ft.restype = c_long
lib.mdPresortGetTotalSacks.argtypes = [c_void_p]
lib.mdPresortGetTotalSacks.restype = c_long
lib.mdPresortGetTotalResidualPieces.argtypes = [c_void_p]
lib.mdPresortGetTotalResidualPieces.restype = c_long
lib.mdPresortGetTotalPiecesAuto5digFSS.argtypes = [c_void_p]
lib.mdPresortGetTotalPiecesAuto5digFSS.restype = c_long
lib.mdPresortGetTotalPieces5digFSS.argtypes = [c_void_p]
lib.mdPresortGetTotalPieces5digFSS.restype = c_long
lib.mdPresortGetRateAuto5digFSS.argtypes = [c_void_p]
lib.mdPresortGetRateAuto5digFSS.restype = c_double
lib.mdPresortGetRate5digFSS.argtypes = [c_void_p]
lib.mdPresortGetRate5digFSS.restype = c_double
lib.mdPresortGetPoundRateAuto5digFSS.argtypes = [c_void_p]
lib.mdPresortGetPoundRateAuto5digFSS.restype = c_double
lib.mdPresortGetPoundRate5digFSS.argtypes = [c_void_p]
lib.mdPresortGetPoundRate5digFSS.restype = c_double
lib.mdPresortGetDropShipZipPlus4.argtypes = [c_void_p]
lib.mdPresortGetDropShipZipPlus4.restype = c_char_p
lib.mdPresortGet.argtypes = [c_void_p, c_char_p]
lib.mdPresortGet.restype = c_char_p
lib.mdPresortGetTotalNonAutoPieces.argtypes = [c_void_p]
lib.mdPresortGetTotalNonAutoPieces.restype = c_long
lib.mdPresortGetPresortedRate.argtypes = [c_void_p]
lib.mdPresortGetPresortedRate.restype = c_double
lib.mdPresortGetResidualRate.argtypes = [c_void_p]
lib.mdPresortGetResidualRate.restype = c_double
lib.mdPresortGetDestTotalPiecesAuto5dig.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesAuto5dig.restype = c_long
lib.mdPresortGetDestTotalPiecesAuto3dig.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesAuto3dig.restype = c_long
lib.mdPresortGetDestTotalPiecesAutoADC.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesAutoADC.restype = c_long
lib.mdPresortGetDestTotalPiecesAutoMADC.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesAutoMADC.restype = c_long
lib.mdPresortGetDestTotalPieces5dig.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPieces5dig.restype = c_long
lib.mdPresortGetDestTotalPieces3dig.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPieces3dig.restype = c_long
lib.mdPresortGetDestTotalPiecesADC.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesADC.restype = c_long
lib.mdPresortGetDestTotalPiecesMADC.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesMADC.restype = c_long
lib.mdPresortGetDestTotalPiecesWS.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesWS.restype = c_long
lib.mdPresortGetDestTotalPiecesHDP.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesHDP.restype = c_long
lib.mdPresortGetDestTotalPiecesHD.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesHD.restype = c_long
lib.mdPresortGetDestTotalPiecesLOT.argtypes = [c_void_p]
lib.mdPresortGetDestTotalPiecesLOT.restype = c_long
lib.mdPresortGetDestRateAuto5dig.argtypes = [c_void_p]
lib.mdPresortGetDestRateAuto5dig.restype = c_double
lib.mdPresortGetDestRateAuto3dig.argtypes = [c_void_p]
lib.mdPresortGetDestRateAuto3dig.restype = c_double
lib.mdPresortGetDestRateAutoADC.argtypes = [c_void_p]
lib.mdPresortGetDestRateAutoADC.restype = c_double
lib.mdPresortGetDestRateAutoMADC.argtypes = [c_void_p]
lib.mdPresortGetDestRateAutoMADC.restype = c_double
lib.mdPresortGetDestRate5dig.argtypes = [c_void_p]
lib.mdPresortGetDestRate5dig.restype = c_double
lib.mdPresortGetDestRate3dig.argtypes = [c_void_p]
lib.mdPresortGetDestRate3dig.restype = c_double
lib.mdPresortGetDestRateADC.argtypes = [c_void_p]
lib.mdPresortGetDestRateADC.restype = c_double
lib.mdPresortGetDestRateMADC.argtypes = [c_void_p]
lib.mdPresortGetDestRateMADC.restype = c_double
lib.mdPresortGetDestRateWS.argtypes = [c_void_p]
lib.mdPresortGetDestRateWS.restype = c_double
lib.mdPresortGetDestRateHDP.argtypes = [c_void_p]
lib.mdPresortGetDestRateHDP.restype = c_double
lib.mdPresortGetDestRateHD.argtypes = [c_void_p]
lib.mdPresortGetDestRateHD.restype = c_double
lib.mdPresortGetDestRateLOT.argtypes = [c_void_p]
lib.mdPresortGetDestRateLOT.restype = c_double
lib.mdPresortGetContainersTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetContainersTotal.restype = c_long
lib.mdPresortGetBundlesTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetBundlesTotal.restype = c_long
lib.mdPresortGetPiecesTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetPiecesTotal.restype = c_long
lib.mdPresortGetChargeTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetChargeTotal.restype = c_double
lib.mdPresortGetDestSCFInfo.argtypes = [c_void_p, c_char_p]
lib.mdPresortGetDestSCFInfo.restype = c_bool
lib.mdPresortGetDestNDCInfo.argtypes = [c_void_p, c_char_p]
lib.mdPresortGetDestNDCInfo.restype = c_bool
lib.mdPresortGetDestDDUInfo.argtypes = [c_void_p, c_char_p]
lib.mdPresortGetDestDDUInfo.restype = c_bool
lib.mdPresortGetDestContainersTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetDestContainersTotal.restype = c_long
lib.mdPresortGetDestBundlesTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetDestBundlesTotal.restype = c_long
lib.mdPresortGetDestPiecesTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetDestPiecesTotal.restype = c_long
lib.mdPresortGetDestChargeTotal.argtypes = [c_void_p, c_int]
lib.mdPresortGetDestChargeTotal.restype = c_double
lib.mdPresortSetTwoFtTrayMaximum.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetTwoFtTrayMaximum.restype = None
lib.mdPresortSetTwoFtTrayMinimum.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetTwoFtTrayMinimum.restype = None
lib.mdPresortSetOneFtTrayMaximum.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetOneFtTrayMaximum.restype = None
lib.mdPresortSetOneFtTrayMinimum.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetOneFtTrayMinimum.restype = None
lib.mdPresortSetBundleMaximum.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetBundleMaximum.restype = None
lib.mdPresortSetPSSCFZipAutomationSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSSCFZipAutomationSort.restype = c_bool
lib.mdPresortSetPSSCFZipNonAutomationSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSSCFZipNonAutomationSort.restype = c_bool
lib.mdPresortSetPSSCFZipECRRTSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSSCFZipECRRTSort.restype = c_bool
lib.mdPresortSetPSSCFZipCoSack.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSSCFZipCoSack.restype = c_bool
lib.mdPresortSetPSNDCZipAutomationSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSNDCZipAutomationSort.restype = c_bool
lib.mdPresortSetPSNDCZipNonAutomationSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSNDCZipNonAutomationSort.restype = c_bool
lib.mdPresortSetPSNDCZipECRRTSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSNDCZipECRRTSort.restype = c_bool
lib.mdPresortSetPSNDCZipCoSack.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSNDCZipCoSack.restype = c_bool
lib.mdPresortSetPSDDUZipECRRTSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSDDUZipECRRTSort.restype = c_bool
lib.mdPresortSetPSPOMZipAutomationSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPOMZipAutomationSort.restype = c_bool
lib.mdPresortSetPSPOMZipNonAutomationSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPOMZipNonAutomationSort.restype = c_bool
lib.mdPresortSetPSPOMZipECRRTSort.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPOMZipECRRTSort.restype = c_bool
lib.mdPresortSetPSPOMZipCoSack.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPOMZipCoSack.restype = c_bool
lib.mdPresortSetPSIncludeResidualPieces.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPSIncludeResidualPieces.restype = None
lib.mdPresortProducePostStatementForResidualPieces.argtypes = [c_void_p]
lib.mdPresortProducePostStatementForResidualPieces.restype = c_bool
lib.mdPresortProducePostStatement.argtypes = [c_void_p]
lib.mdPresortProducePostStatement.restype = None
lib.mdPresortSetPostStatementToSelect.argtypes = [c_void_p, c_bool]
lib.mdPresortSetPostStatementToSelect.restype = None
lib.mdPresortSetIgnoreAspectRatio.argtypes = [c_void_p, c_bool]
lib.mdPresortSetIgnoreAspectRatio.restype = None
lib.mdPresortSetProduceIMBCode.argtypes = [c_void_p, c_bool]
lib.mdPresortSetProduceIMBCode.restype = None
lib.mdPresortGetIMBNumericCode.argtypes = [c_void_p]
lib.mdPresortGetIMBNumericCode.restype = c_char_p
lib.mdPresortGetIMBAlphaCode.argtypes = [c_void_p]
lib.mdPresortGetIMBAlphaCode.restype = c_char_p
lib.mdPresortGetIMBSerialNumber.argtypes = [c_void_p]
lib.mdPresortGetIMBSerialNumber.restype = c_char_p
lib.mdPresortSetPSMailingAgentCRID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSMailingAgentCRID.restype = None
lib.mdPresortSetContainerSequenceNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetContainerSequenceNumber.restype = c_bool
lib.mdPresortSetPSNonProfitAuthNumberMO.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSNonProfitAuthNumberMO.restype = None
lib.mdPresortSetProduceDropShipForms.argtypes = [c_void_p, c_bool]
lib.mdPresortSetProduceDropShipForms.restype = c_bool
lib.mdPresortSetExpandedReportName.argtypes = [c_void_p, c_bool]
lib.mdPresortSetExpandedReportName.restype = None
lib.mdPresortProduceMailDatFiles.argtypes = [c_void_p, c_char_p, c_char_p]
lib.mdPresortProduceMailDatFiles.restype = c_bool
lib.mdPresortSetPSPostOfficeOfMailingPlus4.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPSPostOfficeOfMailingPlus4.restype = None
lib.mdPresortSetMDACSKeyLineData.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDACSKeyLineData.restype = None
lib.mdPresortSetMDMachineID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMachineID.restype = None
lib.mdPresortSetMDJobID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDJobID.restype = None
lib.mdPresortSetMDHDRIDEAllianceVersion.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRIDEAllianceVersion.restype = None
lib.mdPresortSetMDHDRHistorySequenceNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRHistorySequenceNumber.restype = None
lib.mdPresortSetMDHDRHistoryStatus.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRHistoryStatus.restype = None
lib.mdPresortSetMDHDRHistoricalJobID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRHistoricalJobID.restype = None
lib.mdPresortSetMDHDRLicensedUsersJobNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRLicensedUsersJobNumber.restype = None
lib.mdPresortSetMDHDRJobNameTitleIssue.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRJobNameTitleIssue.restype = None
lib.mdPresortSetMDHDRFileSource.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRFileSource.restype = None
lib.mdPresortSetMDHDRUserLicenseCode.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRUserLicenseCode.restype = None
lib.mdPresortSetMDHDRMailDatSoftwareVendorName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRMailDatSoftwareVendorName.restype = None
lib.mdPresortSetMDHDRMailDatSoftwareProductsName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRMailDatSoftwareProductsName.restype = None
lib.mdPresortSetMDHDRMailDatSoftwareVersion.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRMailDatSoftwareVersion.restype = None
lib.mdPresortSetMDHDRMailDatSoftwareVendorEmail.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRMailDatSoftwareVendorEmail.restype = None
lib.mdPresortSetMDHDReDocSenderCRID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDReDocSenderCRID.restype = None
lib.mdPresortSetMDSEGVerificationFacilityName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDSEGVerificationFacilityName.restype = None
lib.mdPresortSetMDSEGVerificationFacilityZipPlus4.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDSEGVerificationFacilityZipPlus4.restype = None
lib.mdPresortSetMDHDROriginalSoftwareVendorName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDROriginalSoftwareVendorName.restype = None
lib.mdPresortSetMDHDROriginalSoftwareProductsName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDROriginalSoftwareProductsName.restype = None
lib.mdPresortSetMDHDROriginalSoftwareVersion.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDROriginalSoftwareVersion.restype = None
lib.mdPresortSetMDHDROriginalSoftwareVendorEmail.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDROriginalSoftwareVendorEmail.restype = None
lib.mdPresortSetMDHDRContactName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRContactName.restype = None
lib.mdPresortSetMDHDRContactPhone.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRContactPhone.restype = None
lib.mdPresortSetMDHDRContactEMail.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDHDRContactEMail.restype = None
lib.mdPresortSetMDCPTMailOwnerID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTMailOwnerID.restype = None
lib.mdPresortSetMDCPTOwnerCRID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTOwnerCRID.restype = None
lib.mdPresortSetMDCPTMailOwnersMailingRefID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTMailOwnersMailingRefID.restype = None
lib.mdPresortSetMDCPTPostalPriceIncID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTPostalPriceIncID.restype = None
lib.mdPresortSetMDCPTPostalPriceIncType.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTPostalPriceIncType.restype = None
lib.mdPresortSetMDCPTStandParcelType.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTStandParcelType.restype = None
lib.mdPresortSetMDCPTStandFlatType.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTStandFlatType.restype = None
lib.mdPresortSetMDCPTUserOptField.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTUserOptField.restype = None
lib.mdPresortSetMDCPTContentOfMail.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTContentOfMail.restype = None
lib.mdPresortSetMDCSMCSAAgreementID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCSMCSAAgreementID.restype = None
lib.mdPresortSetMDSEGDescription.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDSEGDescription.restype = None
lib.mdPresortSetMDCPTComDescription.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCPTComDescription.restype = None
lib.mdPresortSetMDMPUName.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPUName.restype = None
lib.mdPresortSetMDMPUDescription.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPUDescription.restype = None
lib.mdPresortSetMDMPADescription.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPADescription.restype = None
lib.mdPresortSetMDMPAMailingAgentMailerID.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPAMailingAgentMailerID.restype = None
lib.mdPresortSetMDMPAMailOwnerPermitNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPAMailOwnerPermitNumber.restype = None
lib.mdPresortSetMDMPAMailOwnerPermitType.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPAMailOwnerPermitType.restype = None
lib.mdPresortSetMDMPAAdditionalPostagePaymentMethod.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPAAdditionalPostagePaymentMethod.restype = None
lib.mdPresortSetMDMPAAdditionalPaymentPermitNumber.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPAAdditionalPaymentPermitNumber.restype = None
lib.mdPresortSetMDMPAAdditionalPaymentPermitZipPlus4.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDMPAAdditionalPaymentPermitZipPlus4.restype = None
lib.mdPresortSetMDCCRCharacteristic.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCCRCharacteristic.restype = None
lib.mdPresortSetMDCCRCharacteristicType.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetMDCCRCharacteristicType.restype = None
lib.mdPresortSetPOMLocaleKey.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetPOMLocaleKey.restype = None
lib.mdPresortSetDDULocaleKey.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetDDULocaleKey.restype = None
lib.mdPresortSetDDUPostalCode.argtypes = [c_void_p, c_char_p]
lib.mdPresortSetDDUPostalCode.restype = None

# mdPresort Enumerations
class ProgramStatus(Enum):
	ErrorNone = 0
	ErrorOther = 1
	ErrorOutOfMemory = 2
	ErrorRequiredFileNotFound = 3
	ErrorFoundOldFile = 4
	ErrorDatabaseExpired = 5
	ErrorLicenseExpired = 6

class AccessType(Enum):
	Local = 0
	Remote = 1

class DiacriticsMode(Enum):
	Auto = 0
	On = 1
	Off = 2

class StandardizeMode(Enum):
	ShortFormat = 0
	LongFormat = 1
	AutoFormat = 2

class SuiteParseMode(Enum):
	ParseSuite = 0
	CombineSuite = 1

class AliasPreserveMode(Enum):
	ConvertAlias = 0
	PreserveAlias = 1

class AutoCompletionMode(Enum):
	AutoCompleteSingleSuite = 0
	AutoCompleteRangedSuite = 1
	AutoCompletePlaceHolderSuite = 2
	AutoCompleteNoSuite = 3

class ResultCdDescOpt(Enum):
	ResultCodeDescriptionLong = 0
	ResultCodeDescriptionShort = 1

class MailboxLookupMode(Enum):
	MailboxNone = 0
	MailboxExpress = 1
	MailboxPremium = 2

class SortationCode(Enum):
	FCM_LTR_AUTO_NONAUTO = 1
	FCM_LTR_AUTO = 2
	FCM_LTR_NONAUTO = 3
	FCM_LTR_NONMACH = 4
	FCM_POSTCARD_AUTO_NONAUTO = 41
	FCM_POSTCARD_AUTO = 42
	FCM_POSTCARD_NONAUTO = 43
	FCM_FLAT_COTRAY = 51
	FCM_FLAT_AUTO = 52
	FCM_FLAT_NONAUTO = 53
	FCM_FLAT_DISABLE_COTRAY = 54
	FCM_FLAT_COTRAY_FSM1000 = 55
	FCM_FLAT_DISABLE_COTRAY_FSM1000 = 56
	FCM_FLAT_AUTO_FSM1000 = 57
	FCM_FLAT_NONAUTO_FSM1000 = 58
	STD_LTR_ECRRT_AUTO_NONAUTO = 101
	STD_LTR_ECRRT_NONAUTO = 102
	STD_LTR_AUTO_NONAUTO = 103
	STD_LTR_AUTO = 104
	STD_LTR_NONAUTO = 105
	STD_LTR_NONMACH = 106
	STD_LTR_ECRRT_NONMACH = 107
	STD_LTR_ECRRT_AUTO = 108
	STD_LTR_ECRRT = 109
	STD_FLAT_ECRRT_COSACK = 151
	STD_FLAT_ECRRT_DISABLE_COSACK = 152
	STD_FLAT_COSACK = 153
	STD_FLAT_DISABLE_COSACK = 154
	STD_FLAT_AUTO = 155
	STD_FLAT_NONAUTO = 156
	STD_FLAT_ECRRT_AUTO = 157
	STD_FLAT_ECRRT_NONAUTO = 158
	STD_FLAT_ECRRT = 159
	RESIDUALS_LEFT = 999

class ACSCode(Enum):
	FCM_ACS_ASR = 1
	FCM_ACS_ASR2 = 2
	FCM_ACS_CSR = 3
	FCM_ACS_CSR2 = 4
	FCM_ACS_ASR_ONECODE = 5
	FCM_ACS_ASR2_ONECODE = 6
	FCM_ACS_CSR_ONECODE = 7
	FCM_ACS_CSR2_ONECODE = 8
	FCM_MANUAL_CORRECTION = 9
	FCM_NO_ADDR_CORRECTION = 10
	FCM_ACS_ASR_TRACE = 11
	FCM_ACS_ASR2_TRACE = 12
	FCM_ACS_CSR_TRACE = 13
	FCM_ACS_CSR2_TRACE = 14
	FCM_ACS_ASR_ONECODE_TRACE = 15
	FCM_ACS_CSR_ONECODE_TRACE = 16
	FCM_MANUAL_CORRECTION_TRACE = 17
	FCM_NO_ADDR_CORRECTION_TRACE = 18
	FCM_ACS_ASR_FS = 19
	FCM_ACS_ASR2_FS = 20
	FCM_ACS_CSR_FS = 21
	FCM_ACS_CSR2_FS = 22
	FCM_FULL_SERVICE_ACS_ASR_FS = 23
	FCM_FULL_SERVICE_ACS_ASR2_FS = 24
	FCM_NO_ADDR_CORRECTION_FS = 25
	FCM_ACS_ASR_FS_TRACE = 26
	FCM_ACS_ASR2_FS_TRACE = 27
	FCM_ACS_CSR_FS_TRACE = 28
	FCM_ACS_CSR2_FS_TRACE = 29
	FCM_FULL_SERVICE_ACS_ASR_FS_TRACE = 30
	FCM_FULL_SERVICE_ACS_ASR2_FS_TRACE = 31
	FCM_NO_ADDR_CORRECTION_FS_TRACE = 32
	FCM_ACS_ASR2_ONECODE_TRACE = 33
	FCM_ACS_CSR2_ONECODE_TRACE = 34
	FCM_FULL_SERVICE_ACS_CSR_FS = 35
	FCM_FULL_SERVICE_ACS_CSR2_FS = 36
	FCM_FULL_SERVICE_ACS_CSR_FS_TRACE = 37
	FCM_FULL_SERVICE_ACS_CSR2_FS_TRACE = 38
	FCM_ACS_RSR2 = 39
	FCM_ACS_TRSR2 = 40
	FCM_ACS_RSR2_ONECODE = 41
	FCM_ACS_TRSR2_ONECODE = 42
	FCM_ACS_RSR2_TRACE = 43
	FCM_ACS_TRSR2_TRACE = 44
	FCM_ACS_RSR2_ONECODE_TRACE = 45
	FCM_ACS_TRSR2_ONECODE_TRACE = 46
	FCM_ACS_RSR2_FS = 47
	FCM_ACS_TRSR2_FS = 48
	FCM_FULL_SERVICE_ACS_RSR2_FS = 49
	FCM_FULL_SERVICE_ACS_TRSR2_FS = 50
	FCM_ACS_RSR2_FS_TRACE = 51
	FCM_ACS_TRSR2_FS_TRACE = 52
	FCM_FULL_SERVICE_ACS_RSR2_FS_TRACE = 53
	FCM_FULL_SERVICE_ACS_TRSR2_FS_TRACE = 54
	FCM_MANUAL_CORRECTION_POLITICALMAIL = 55
	FCM_NO_ADDR_CORRECTION_POLITICALMAIL = 56
	FCM_MANUAL_CORRECTION_TRACE_POLITICALMAIL = 57
	FCM_NO_ADDR_CORRECTION_TRACE_POLITICALMAIL = 58
	FCM_NO_ADDR_CORRECTION_FS_POLITICALMAIL = 59
	FCM_NO_ADDR_CORRECTION_TRACE_FS_POLITICALMAIL = 60
	FCM_ACS_ASR_ONECODE_POLITICALMAIL = 61
	FCM_ACS_ASR_ONECODE_TRACE_POLITICALMAIL = 62
	FCM_ACS_ASR2_ONECODE_POLITICALMAIL = 63
	FCM_ACS_ASR2_ONECODE_TRACE_POLITICALMAIL = 64
	FCM_ACS_CSR_ONECODE_POLITICALMAIL = 65
	FCM_ACS_CSR_ONECODE_TRACE_POLITICALMAIL = 66
	FCM_FULL_SERVICE_ACS_ASR_POLITICALMAIL = 67
	FCM_FULL_SERVICE_ACS_ASR_TRACE_POLITICALMAIL = 68
	FCM_FULL_SERVICE_ACS_ASR2_POLITICALMAIL = 69
	FCM_FULL_SERVICE_ACS_ASR2_TRACE_POLITICALMAIL = 70
	FCM_FULL_SERVICE_ACS_CSR_POLITICALMAIL = 71
	FCM_FULL_SERVICE_ACS_CSR_TRACE_POLITICALMAIL = 72
	FCM_FULL_SERVICE_ACS_RSR2_POLITICALMAIL = 73
	FCM_FULL_SERVICE_ACS_RSR2_TRACE_POLITICALMAIL = 74
	FCM_NO_ADDR_CORRECTION_TRACE_BALLOTMAIL = 75
	FCM_NO_ADDR_CORRECTION_TRACE_FS_BALLOTMAIL = 76
	FCM_MANUAL_CORRECTION_TRACE_BALLOTMAIL = 77
	FCM_ACS_ASR_ONECODE_TRACE_BALLOTMAIL = 78
	FCM_ACS_ASR2_ONECODE_TRACE_BALLOTMAIL = 79
	FCM_ACS_CSR_ONECODE_TRACE_BALLOTMAIL = 80
	FCM_ACS_RSR2_ONECODE_TRACE_BALLOTMAIL = 81
	FCM_FULL_SERVICE_ACS_ASR_TRACE_FS_BALLOTMAIL = 82
	FCM_FULL_SERVICE_ACS_ASR2_TRACE_FS_BALLOTMAIL = 83
	FCM_FULL_SERVICE_ACS_CSR_TRACE_FS_BALLOTMAIL = 84
	FCM_FULL_SERVICE_ACS_RSR2_TRACE_FS_BALLOTMAIL = 85
	STD_ACS_ASR = 101
	STD_ACS_CSR = 102
	STD_ACS_ASR_ONECODE = 103
	STD_ACS_CSR_ONECODE = 104
	STD_MANUAL_CORRECTION = 105
	STD_NO_ADDR_CORRECTION = 106
	STD_ACS_ASR_TRACE = 107
	STD_ACS_CSR_TRACE = 108
	STD_ACS_ASR_ONECODE_TRACE = 109
	STD_ACS_CSR_ONECODE_TRACE = 110
	STD_MANUAL_CORRECTION_TRACE = 111
	STD_NO_ADDR_CORRECTION_TRACE = 112
	STD_ACS_ASR_FS = 113
	STD_ACS_CSR_FS = 114
	STD_FULL_SERVICE_ACS_ASR_FS = 115
	STD_FULL_SERVICE_ACS_CSR_FS = 116
	STD_NO_ADDR_CORRECTION_FS = 117
	STD_ACS_ASR_FS_TRACE = 118
	STD_ACS_CSR_FS_TRACE = 119
	STD_FULL_SERVICE_ACS_ASR_FS_TRACE = 120
	STD_FULL_SERVICE_ACS_CSR_FS_TRACE = 121
	STD_NO_ADDR_CORRECTION_FS_TRACE = 122
	STD_ACS_ASR2 = 123
	STD_ACS_RSR2 = 124
	STD_ACS_ASR2_ONECODE = 125
	STD_ACS_RSR2_ONECODE = 126
	STD_ACS_ASR2_TRACE = 127
	STD_ACS_RSR2_TRACE = 128
	STD_ACS_ASR2_ONECODE_TRACE = 129
	STD_ACS_RSR2_ONECODE_TRACE = 130
	STD_ACS_ASR2_FS = 131
	STD_ACS_RSR2_FS = 132
	STD_FULL_SERVICE_ACS_ASR2_FS = 133
	STD_FULL_SERVICE_ACS_RSR2_FS = 134
	STD_ACS_ASR2_FS_TRACE = 135
	STD_ACS_RSR2_FS_TRACE = 136
	STD_FULL_SERVICE_ACS_ASR2_FS_TRACE = 137
	STD_FULL_SERVICE_ACS_RSR2_FS_TRACE = 138
	STD_ACS_CSR2 = 139
	STD_ACS_CSR2_TRACE = 140
	STD_ACS_CSR2_ONECODE = 141
	STD_ACS_CSR2_ONECODE_TRACE = 142
	STD_ACS_CSR2_FS = 143
	STD_ACS_CSR2_FS_TRACE = 144
	STD_FULL_SERVICE_ACS_CSR2_FS = 145
	STD_FULL_SERVICE_ACS_CSR2_FS_TRACE = 146
	STD_MANUAL_CORRECTION_POLITICALMAIL = 147
	STD_NO_ADDR_CORRECTION_POLITICALMAIL = 148
	STD_MANUAL_CORRECTION_TRACE_POLITICALMAIL = 149
	STD_NO_ADDR_CORRECTION_TRACE_POLITICALMAIL = 150
	STD_NO_ADDR_CORRECTION_FS_POLITICALMAIL = 151
	STD_NO_ADDR_CORRECTION_TRACE_FS_POLITICALMAIL = 152
	STD_ACS_CSR_ONECODE_POLITICALMAIL = 153
	STD_ACS_CSR_ONECODE_TRACE_POLITICALMAIL = 154
	STD_FULL_SERVICE_ACS_CSR_POLITICALMAIL = 155
	STD_FULL_SERVICE_ACS_CSR_TRACE_POLITICALMAIL = 156
	STD_ACS_CSR_POLITICALMAIL = 157
	STD_ACS_CSR_TRACE_POLITICALMAIL = 158
	STD_NO_ADDR_CORRECTION_TRACE_BALLOTMAIL = 159
	STD_NO_ADDR_CORRECTION_TRACE_FS_BALLOTMAIL = 160
	STD_MANUAL_CORRECTION_TRACE_BALLOTMAIL = 161
	STD_ACS_ASR_ONECODE_TRACE_BALLOTMAIL = 162
	STD_ACS_ASR2_ONECODE_TRACE_BALLOTMAIL = 163
	STD_ACS_CSR_ONECODE_TRACE_BALLOTMAIL = 164
	STD_ACS_RSR2_ONECODE_TRACE_BALLOTMAIL = 165
	STD_FULL_SERVICE_ACS_ASR_TRACE_FS_BALLOTMAIL = 166
	STD_FULL_SERVICE_ACS_ASR2_TRACE_FS_BALLOTMAIL = 167
	STD_FULL_SERVICE_ACS_CSR_TRACE_FS_BALLOTMAIL = 168
	STD_FULL_SERVICE_ACS_RSR2_TRACE_FS_BALLOTMAIL = 169
	STD_ACS_CSR_TRACE_BALLOTMAIL = 170
	STD_ACS_RSR2_TRACE_BALLOTMAIL = 171
	RETURN_BALLOT_FCM_MAIL_REPLY_ENVELOPES_TRACE = 300
	RETURN_BALLOT_BUSINESS_REPLY_MAIL_BY_ZIP_ENVELOPES_TRACE = 301
	RETURN_BALLOT_PERMIT_REPLY_MAIL_BY_ZIP_ENVELOPES_TRACE = 302
	RETURN_BALLOT_UOCAVA_TRACE = 303
	ACS_LAST = 999

class mdPresort(object):
	def __init__(self):
		self.I = lib.mdPresortCreate()

	def __del__(self):
		lib.mdPresortDestroy(self.I)

	def InitializeDataFiles(self):
		return ProgramStatus(lib.mdPresortInitializeDataFiles(self.I))

	def SetLicenseString(self, p1):
		return lib.mdPresortSetLicenseString(self.I, p1.encode('utf-8'))

	def SetPathToPresortDataFiles(self, p1):
		lib.mdPresortSetPathToPresortDataFiles(self.I, p1.encode('utf-8'))

	def GetInitializeErrorString(self):
		return lib.mdPresortGetInitializeErrorString(self.I).decode('utf-8')

	def GetParametersErrorString(self):
		return lib.mdPresortGetParametersErrorString(self.I).decode('utf-8')

	def GetBuildNumber(self):
		return lib.mdPresortGetBuildNumber(self.I).decode('utf-8')

	def GetLicenseStringExpirationDate(self):
		return lib.mdPresortGetLicenseStringExpirationDate(self.I).decode('utf-8')

	def GetDatabaseExpirationDate(self):
		return lib.mdPresortGetDatabaseExpirationDate(self.I).decode('utf-8')

	def GetDatabaseDate(self):
		return lib.mdPresortGetDatabaseDate(self.I).decode('utf-8')

	def SetZip(self, p1):
		lib.mdPresortSetZip(self.I, p1.encode('utf-8'))

	def SetPlus4(self, p1):
		lib.mdPresortSetPlus4(self.I, p1.encode('utf-8'))

	def SetCarrierRoute(self, p1):
		lib.mdPresortSetCarrierRoute(self.I, p1.encode('utf-8'))

	def SetWalkSequence(self, p1):
		lib.mdPresortSetWalkSequence(self.I, p1.encode('utf-8'))

	def SetDeliveryPointCode(self, p1):
		lib.mdPresortSetDeliveryPointCode(self.I, p1.encode('utf-8'))

	def SetBusinessResidentialIndicator(self, p1):
		lib.mdPresortSetBusinessResidentialIndicator(self.I, p1.encode('utf-8'))

	def SetPSPermitImprint(self, p1):
		lib.mdPresortSetPSPermitImprint(self.I, p1)

	def SetLOTNumber(self, p1):
		lib.mdPresortSetLOTNumber(self.I, p1.encode('utf-8'))

	def SetLOTOrder(self, p1):
		lib.mdPresortSetLOTOrder(self.I, p1.encode('utf-8'))

	def SetPSASE(self, p1):
		lib.mdPresortSetPSASE(self.I, p1)

	def SetPSFASTforward(self, p1):
		lib.mdPresortSetPSFASTforward(self.I, p1)

	def SetPSNCOA(self, p1):
		lib.mdPresortSetPSNCOA(self.I, p1)

	def SetPSACS(self, p1):
		lib.mdPresortSetPSACS(self.I, p1)

	def SetPSAltMethod(self, p1):
		lib.mdPresortSetPSAltMethod(self.I, p1)

	def SetPSMultiple(self, p1):
		lib.mdPresortSetPSMultiple(self.I, p1)

	def SetPSOneCode(self, p1):
		lib.mdPresortSetPSOneCode(self.I, p1)

	def SetPSAltAddFmt(self, p1):
		lib.mdPresortSetPSAltAddFmt(self.I, p1)

	def SetPSCASSDate(self, p1):
		lib.mdPresortSetPSCASSDate(self.I, p1.encode('utf-8'))

	def SetPSPrecanceledStamp(self, p1):
		lib.mdPresortSetPSPrecanceledStamp(self.I, p1)

	def SetPSPermitHolderName(self, p1):
		lib.mdPresortSetPSPermitHolderName(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderCompany(self, p1):
		lib.mdPresortSetPSPermitHolderCompany(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderStreet(self, p1):
		lib.mdPresortSetPSPermitHolderStreet(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderCity(self, p1):
		lib.mdPresortSetPSPermitHolderCity(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderState(self, p1):
		lib.mdPresortSetPSPermitHolderState(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderPhone(self, p1):
		lib.mdPresortSetPSPermitHolderPhone(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderEmail(self, p1):
		lib.mdPresortSetPSPermitHolderEmail(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderListName(self, p1):
		lib.mdPresortSetPSPermitHolderListName(self.I, p1.encode('utf-8'))

	def SetPSPermitHolderZIP(self, p1):
		lib.mdPresortSetPSPermitHolderZIP(self.I, p1.encode('utf-8'))

	def SetPSPostOfficeOfMailingCity(self, p1):
		lib.mdPresortSetPSPostOfficeOfMailingCity(self.I, p1.encode('utf-8'))

	def SetPSPostOfficeOfMailingState(self, p1):
		lib.mdPresortSetPSPostOfficeOfMailingState(self.I, p1.encode('utf-8'))

	def SetPSPostOfficeOfMailingZIP(self, p1):
		lib.mdPresortSetPSPostOfficeOfMailingZIP(self.I, p1.encode('utf-8'))

	def SetPSMailingAgentName(self, p1):
		lib.mdPresortSetPSMailingAgentName(self.I, p1.encode('utf-8'))

	def SetPSMailingAgentCompany(self, p1):
		lib.mdPresortSetPSMailingAgentCompany(self.I, p1.encode('utf-8'))

	def SetPSMailingAgentStreet(self, p1):
		lib.mdPresortSetPSMailingAgentStreet(self.I, p1.encode('utf-8'))

	def SetPSMailingAgentCity(self, p1):
		lib.mdPresortSetPSMailingAgentCity(self.I, p1.encode('utf-8'))

	def SetPSMailingAgentState(self, p1):
		lib.mdPresortSetPSMailingAgentState(self.I, p1.encode('utf-8'))

	def SetPSMailingAgentPhone(self, p1):
		lib.mdPresortSetPSMailingAgentPhone(self.I, p1.encode('utf-8'))

	def SetPSMailingAgentZIP(self, p1):
		lib.mdPresortSetPSMailingAgentZIP(self.I, p1.encode('utf-8'))

	def SetPSIndividualOrOrganizationName(self, p1):
		lib.mdPresortSetPSIndividualOrOrganizationName(self.I, p1.encode('utf-8'))

	def SetPSIndividualOrOrganizationCompany(self, p1):
		lib.mdPresortSetPSIndividualOrOrganizationCompany(self.I, p1.encode('utf-8'))

	def SetPSIndividualOrOrganizationStreet(self, p1):
		lib.mdPresortSetPSIndividualOrOrganizationStreet(self.I, p1.encode('utf-8'))

	def SetPSIndividualOrOrganizationCity(self, p1):
		lib.mdPresortSetPSIndividualOrOrganizationCity(self.I, p1.encode('utf-8'))

	def SetPSIndividualOrOrganizationState(self, p1):
		lib.mdPresortSetPSIndividualOrOrganizationState(self.I, p1.encode('utf-8'))

	def SetPSIndividualOrOrganizationZIP(self, p1):
		lib.mdPresortSetPSIndividualOrOrganizationZIP(self.I, p1.encode('utf-8'))

	def SetPSIndividualOrOrganizationCRID(self, p1):
		lib.mdPresortSetPSIndividualOrOrganizationCRID(self.I, p1.encode('utf-8'))

	def SetContinueContainerNumber(self, p1):
		lib.mdPresortSetContinueContainerNumber(self.I, p1)

	def SetPOMasNDC(self, p1):
		return lib.mdPresortSetPOMasNDC(self.I, p1)

	def SetPOMasSCF(self, p1):
		return lib.mdPresortSetPOMasSCF(self.I, p1)

	def SetPSStatementSeqNumber(self, val):
		lib.mdPresortSetPSStatementSeqNumber(self.I, val.encode('utf-8'))

	def SetPSMailingDate(self, val):
		lib.mdPresortSetPSMailingDate(self.I, val.encode('utf-8'))

	def SetPSFedAgencyCode(self, val):
		lib.mdPresortSetPSFedAgencyCode(self.I, val.encode('utf-8'))

	def SetPSCustomerNumber(self, val):
		lib.mdPresortSetPSCustomerNumber(self.I, val.encode('utf-8'))

	def SetPSCAPSNumber(self, val):
		lib.mdPresortSetPSCAPSNumber(self.I, val.encode('utf-8'))

	def SetPermitNumber(self, p1):
		lib.mdPresortSetPermitNumber(self.I, p1.encode('utf-8'))

	def SetPSNonProfitAuthNumber(self, p1):
		lib.mdPresortSetPSNonProfitAuthNumber(self.I, p1.encode('utf-8'))

	def SetMailClass(self, p1):
		lib.mdPresortSetMailClass(self.I, p1.encode('utf-8'))

	def SetPieceType(self, p1):
		lib.mdPresortSetPieceType(self.I, p1.encode('utf-8'))

	def SetSortType(self, p1):
		lib.mdPresortSetSortType(self.I, p1.encode('utf-8'))

	def SetFCM_CRRT(self, p1):
		lib.mdPresortSetFCM_CRRT(self.I, p1)

	def SetFCM_CRRT_5(self, p1):
		lib.mdPresortSetFCM_CRRT_5(self.I, p1)

	def SetFCM_CRRT_3(self, p1):
		lib.mdPresortSetFCM_CRRT_3(self.I, p1)

	def SetFCM_5dg_Scheme(self, p1):
		lib.mdPresortSetFCM_5dg_Scheme(self.I, p1)

	def SetFCM_Auto_5dg(self, p1):
		lib.mdPresortSetFCM_Auto_5dg(self.I, p1)

	def SetFCM_NonAuto_5dg(self, p1):
		lib.mdPresortSetFCM_NonAuto_5dg(self.I, p1)

	def SetSTD_Auto_5dg_Scheme(self, p1):
		lib.mdPresortSetSTD_Auto_5dg_Scheme(self.I, p1)

	def SetSTD_Auto_5dg(self, p1):
		lib.mdPresortSetSTD_Auto_5dg(self.I, p1)

	def SetNonMachinableType(self, p1):
		lib.mdPresortSetNonMachinableType(self.I, p1)

	def SetRecordID(self, p1):
		lib.mdPresortSetRecordID(self.I, p1.encode('utf-8'))

	def SetSackWeight(self, p1):
		lib.mdPresortSetSackWeight(self.I, p1.encode('utf-8'))

	def SetPieceLength(self, p1):
		lib.mdPresortSetPieceLength(self.I, p1.encode('utf-8'))

	def SetPieceHeight(self, p1):
		lib.mdPresortSetPieceHeight(self.I, p1.encode('utf-8'))

	def SetPieceThickness(self, p1):
		lib.mdPresortSetPieceThickness(self.I, p1.encode('utf-8'))

	def SetPieceWeight(self, p1):
		lib.mdPresortSetPieceWeight(self.I, p1.encode('utf-8'))

	def SetPackageSize(self, p1):
		lib.mdPresortSetPackageSize(self.I, p1.encode('utf-8'))

	def SetEnableCoTray(self, p1):
		lib.mdPresortSetEnableCoTray(self.I, p1)

	def SetUseFSM1000(self, p1):
		lib.mdPresortSetUseFSM1000(self.I, p1)

	def SetSTDNonProfit(self, p1):
		lib.mdPresortSetSTDNonProfit(self.I, p1)

	def SetIgnoreDSF(self, val):
		lib.mdPresortSetIgnoreDSF(self.I, val)

	def SetMailersID(self, p1):
		lib.mdPresortSetMailersID(self.I, p1.encode('utf-8'))

	def SetPSPrecanceledStampValue(self, p1):
		return lib.mdPresortSetPSPrecanceledStampValue(self.I, p1.encode('utf-8'))

	def SetPSPresortResidualPieces(self, p1):
		lib.mdPresortSetPSPresortResidualPieces(self.I, p1)

	def SetIMBSerialNumber(self, p1):
		lib.mdPresortSetIMBSerialNumber(self.I, p1.encode('utf-8'))

	def SetPSPolicitalCampaignMailing(self, p1):
		lib.mdPresortSetPSPolicitalCampaignMailing(self.I, p1)

	def SetPSOfficialElectionMail(self, p1):
		lib.mdPresortSetPSOfficialElectionMail(self.I, p1)

	def ProduceReports(self, p1, p2):
		return lib.mdPresortProduceReports(self.I, p1.encode('utf-8'), p2.encode('utf-8'))

	def SetSCFCity(self, p1):
		lib.mdPresortSetSCFCity(self.I, p1.encode('utf-8'))

	def SetSCFState(self, p1):
		lib.mdPresortSetSCFState(self.I, p1.encode('utf-8'))

	def SetSCFZip(self, p1):
		lib.mdPresortSetSCFZip(self.I, p1.encode('utf-8'))

	def AddSCF(self):
		return lib.mdPresortAddSCF(self.I)

	def SetNDCCity(self, p1):
		lib.mdPresortSetNDCCity(self.I, p1.encode('utf-8'))

	def SetNDCState(self, p1):
		lib.mdPresortSetNDCState(self.I, p1.encode('utf-8'))

	def SetNDCZip(self, p1):
		lib.mdPresortSetNDCZip(self.I, p1.encode('utf-8'))

	def AddNDC(self):
		return lib.mdPresortAddNDC(self.I)

	def SetDDUCity(self, p1):
		lib.mdPresortSetDDUCity(self.I, p1.encode('utf-8'))

	def SetDDUState(self, p1):
		lib.mdPresortSetDDUState(self.I, p1.encode('utf-8'))

	def SetDDUZip(self, p1):
		lib.mdPresortSetDDUZip(self.I, p1.encode('utf-8'))

	def SetDDUMoreZip(self, p1):
		lib.mdPresortSetDDUMoreZip(self.I, p1.encode('utf-8'))

	def AddDDU(self):
		return lib.mdPresortAddDDU(self.I)

	def SetTTBorder(self, p1):
		lib.mdPresortSetTTBorder(self.I, p1)

	def SetTTNumberofPieces(self, p1):
		lib.mdPresortSetTTNumberofPieces(self.I, p1)

	def SetTTContainerNumber(self, val):
		lib.mdPresortSetTTContainerNumber(self.I, val)

	def SetTTContainerSize(self, val):
		lib.mdPresortSetTTContainerSize(self.I, val)

	def SetTTOther(self, val2):
		lib.mdPresortSetTTOther(self.I, val2.encode('utf-8'))

	def SetTTParameterPositionX(self, val):
		lib.mdPresortSetTTParameterPositionX(self.I, val.encode('utf-8'))

	def SetTTParameterPositionY(self, val):
		lib.mdPresortSetTTParameterPositionY(self.I, val.encode('utf-8'))

	def SetTTParameterWidth(self, val):
		lib.mdPresortSetTTParameterWidth(self.I, val.encode('utf-8'))

	def SetTTParameterHeight(self, val):
		lib.mdPresortSetTTParameterHeight(self.I, val.encode('utf-8'))

	def DoPresort(self):
		return lib.mdPresortDoPresort(self.I)

	def SetDisableDetailedOutput(self, p1):
		lib.mdPresortSetDisableDetailedOutput(self.I, p1)

	def SetPreSortSettings(self, p1):
		return lib.mdPresortSetPreSortSettings(self.I, p1).decode('utf-8')

	def SetACSCodeSettings(self, p1):
		return lib.mdPresortSetACSCodeSettings(self.I, p1)

	def UpdateParameters(self):
		return lib.mdPresortUpdateParameters(self.I)

	def AddRecord(self):
		return lib.mdPresortAddRecord(self.I)

	def GetRecord(self, p1):
		return lib.mdPresortGetRecord(self.I, p1.encode('utf-8'))

	def GetFirstRecord(self):
		return lib.mdPresortGetFirstRecord(self.I)

	def GetNextRecord(self):
		return lib.mdPresortGetNextRecord(self.I)

	def GetMailPieceRate(self):
		return lib.mdPresortGetMailPieceRate(self.I).decode('utf-8')

	def GetZipAsString(self):
		return lib.mdPresortGetZipAsString(self.I).decode('utf-8')

	def GetTrayNumber(self):
		return lib.mdPresortGetTrayNumber(self.I)

	def GetSequenceNumber(self):
		return lib.mdPresortGetSequenceNumber(self.I)

	def GetRecordID(self):
		return lib.mdPresortGetRecordID(self.I).decode('utf-8')

	def GetEndorsementLine(self):
		return lib.mdPresortGetEndorsementLine(self.I).decode('utf-8')

	def GetRateCode(self):
		return lib.mdPresortGetRateCode(self.I).decode('utf-8')

	def GetMailJob(self):
		return lib.mdPresortGetMailJob(self.I).decode('utf-8')

	def GetSortLevel(self):
		return lib.mdPresortGetSortLevel(self.I).decode('utf-8')

	def GetLabelLine1(self):
		return lib.mdPresortGetLabelLine1(self.I).decode('utf-8')

	def GetLabelLine2(self):
		return lib.mdPresortGetLabelLine2(self.I).decode('utf-8')

	def GetLabelList(self):
		return lib.mdPresortGetLabelList(self.I).decode('utf-8')

	def GetCINCode(self):
		return lib.mdPresortGetCINCode(self.I).decode('utf-8')

	def GetTrayProcessingCode(self):
		return lib.mdPresortGetTrayProcessingCode(self.I).decode('utf-8')

	def GetMailSplitCode(self):
		return lib.mdPresortGetMailSplitCode(self.I).decode('utf-8')

	def GetTrayZipCode(self):
		return lib.mdPresortGetTrayZipCode(self.I).decode('utf-8')

	def GetTrayType(self):
		return lib.mdPresortGetTrayType(self.I).decode('utf-8')

	def GetBundleZipCode(self):
		return lib.mdPresortGetBundleZipCode(self.I).decode('utf-8')

	def GetBundleSortLevel(self):
		return lib.mdPresortGetBundleSortLevel(self.I).decode('utf-8')

	def GetBundleNumber(self):
		return lib.mdPresortGetBundleNumber(self.I)

	def GetBundleIndicator(self):
		return lib.mdPresortGetBundleIndicator(self.I).decode('utf-8')

	def GetZipCodeInScheme(self):
		return lib.mdPresortGetZipCodeInScheme(self.I).decode('utf-8')

	def GetBarcodeID(self):
		return lib.mdPresortGetBarcodeID(self.I).decode('utf-8')

	def GetServiceTypeID(self):
		return lib.mdPresortGetServiceTypeID(self.I).decode('utf-8')

	def SetProducePallets(self, p1):
		return lib.mdPresortSetProducePallets(self.I, p1)

	def SetIgnoreMinSCFPallet(self, val):
		lib.mdPresortSetIgnoreMinSCFPallet(self.I, val)

	def GetPalletsTotal(self):
		return lib.mdPresortGetPalletsTotal(self.I)

	def GetPalletZipCode(self):
		return lib.mdPresortGetPalletZipCode(self.I).decode('utf-8')

	def GetPalletLabelLine1(self):
		return lib.mdPresortGetPalletLabelLine1(self.I).decode('utf-8')

	def GetPalletLabelLine2(self):
		return lib.mdPresortGetPalletLabelLine2(self.I).decode('utf-8')

	def GetPalletSortLevel(self):
		return lib.mdPresortGetPalletSortLevel(self.I).decode('utf-8')

	def GetPalletNumber(self):
		return lib.mdPresortGetPalletNumber(self.I)

	def GetTotalPiecesAuto5dig(self):
		return lib.mdPresortGetTotalPiecesAuto5dig(self.I)

	def GetTotalPiecesAuto3dig(self):
		return lib.mdPresortGetTotalPiecesAuto3dig(self.I)

	def GetTotalPiecesAutoADC(self):
		return lib.mdPresortGetTotalPiecesAutoADC(self.I)

	def GetTotalPiecesAutoMADC(self):
		return lib.mdPresortGetTotalPiecesAutoMADC(self.I)

	def GetTotalPieces5dig(self):
		return lib.mdPresortGetTotalPieces5dig(self.I)

	def GetTotalPieces3dig(self):
		return lib.mdPresortGetTotalPieces3dig(self.I)

	def GetTotalPiecesADC(self):
		return lib.mdPresortGetTotalPiecesADC(self.I)

	def GetTotalPiecesMADC(self):
		return lib.mdPresortGetTotalPiecesMADC(self.I)

	def GetTotalPiecesWS(self):
		return lib.mdPresortGetTotalPiecesWS(self.I)

	def GetTotalPiecesHDP(self):
		return lib.mdPresortGetTotalPiecesHDP(self.I)

	def GetTotalPiecesHD(self):
		return lib.mdPresortGetTotalPiecesHD(self.I)

	def GetTotalPiecesLOT(self):
		return lib.mdPresortGetTotalPiecesLOT(self.I)

	def GetRateAuto5dig(self):
		return lib.mdPresortGetRateAuto5dig(self.I)

	def GetRateAuto3dig(self):
		return lib.mdPresortGetRateAuto3dig(self.I)

	def GetRateAutoADC(self):
		return lib.mdPresortGetRateAutoADC(self.I)

	def GetRateAutoMADC(self):
		return lib.mdPresortGetRateAutoMADC(self.I)

	def GetRate5dig(self):
		return lib.mdPresortGetRate5dig(self.I)

	def GetRate3dig(self):
		return lib.mdPresortGetRate3dig(self.I)

	def GetRateADC(self):
		return lib.mdPresortGetRateADC(self.I)

	def GetRateMADC(self):
		return lib.mdPresortGetRateMADC(self.I)

	def GetRateWS(self):
		return lib.mdPresortGetRateWS(self.I)

	def GetRateHDP(self):
		return lib.mdPresortGetRateHDP(self.I)

	def GetRateHD(self):
		return lib.mdPresortGetRateHD(self.I)

	def GetRateLOT(self):
		return lib.mdPresortGetRateLOT(self.I)

	def GetPoundRate(self):
		return lib.mdPresortGetPoundRate(self.I)

	def GetPoundRateWS(self):
		return lib.mdPresortGetPoundRateWS(self.I)

	def GetPoundRateHDP(self):
		return lib.mdPresortGetPoundRateHDP(self.I)

	def GetPoundRateHD(self):
		return lib.mdPresortGetPoundRateHD(self.I)

	def GetPoundRateLOT(self):
		return lib.mdPresortGetPoundRateLOT(self.I)

	def GetTotalTrays1Ft(self):
		return lib.mdPresortGetTotalTrays1Ft(self.I)

	def GetTotalTrays2Ft(self):
		return lib.mdPresortGetTotalTrays2Ft(self.I)

	def GetTotalSacks(self):
		return lib.mdPresortGetTotalSacks(self.I)

	def GetTotalResidualPieces(self):
		return lib.mdPresortGetTotalResidualPieces(self.I)

	def GetTotalPiecesAuto5digFSS(self):
		return lib.mdPresortGetTotalPiecesAuto5digFSS(self.I)

	def GetTotalPieces5digFSS(self):
		return lib.mdPresortGetTotalPieces5digFSS(self.I)

	def GetRateAuto5digFSS(self):
		return lib.mdPresortGetRateAuto5digFSS(self.I)

	def GetRate5digFSS(self):
		return lib.mdPresortGetRate5digFSS(self.I)

	def GetPoundRateAuto5digFSS(self):
		return lib.mdPresortGetPoundRateAuto5digFSS(self.I)

	def GetPoundRate5digFSS(self):
		return lib.mdPresortGetPoundRate5digFSS(self.I)

	def GetDropShipZipPlus4(self):
		return lib.mdPresortGetDropShipZipPlus4(self.I).decode('utf-8')

	def Get(self, p1):
		return lib.mdPresortGet(self.I, p1.encode('utf-8')).decode('utf-8')

	def GetTotalNonAutoPieces(self):
		return lib.mdPresortGetTotalNonAutoPieces(self.I)

	def GetPresortedRate(self):
		return lib.mdPresortGetPresortedRate(self.I)

	def GetResidualRate(self):
		return lib.mdPresortGetResidualRate(self.I)

	def GetDestTotalPiecesAuto5dig(self):
		return lib.mdPresortGetDestTotalPiecesAuto5dig(self.I)

	def GetDestTotalPiecesAuto3dig(self):
		return lib.mdPresortGetDestTotalPiecesAuto3dig(self.I)

	def GetDestTotalPiecesAutoADC(self):
		return lib.mdPresortGetDestTotalPiecesAutoADC(self.I)

	def GetDestTotalPiecesAutoMADC(self):
		return lib.mdPresortGetDestTotalPiecesAutoMADC(self.I)

	def GetDestTotalPieces5dig(self):
		return lib.mdPresortGetDestTotalPieces5dig(self.I)

	def GetDestTotalPieces3dig(self):
		return lib.mdPresortGetDestTotalPieces3dig(self.I)

	def GetDestTotalPiecesADC(self):
		return lib.mdPresortGetDestTotalPiecesADC(self.I)

	def GetDestTotalPiecesMADC(self):
		return lib.mdPresortGetDestTotalPiecesMADC(self.I)

	def GetDestTotalPiecesWS(self):
		return lib.mdPresortGetDestTotalPiecesWS(self.I)

	def GetDestTotalPiecesHDP(self):
		return lib.mdPresortGetDestTotalPiecesHDP(self.I)

	def GetDestTotalPiecesHD(self):
		return lib.mdPresortGetDestTotalPiecesHD(self.I)

	def GetDestTotalPiecesLOT(self):
		return lib.mdPresortGetDestTotalPiecesLOT(self.I)

	def GetDestRateAuto5dig(self):
		return lib.mdPresortGetDestRateAuto5dig(self.I)

	def GetDestRateAuto3dig(self):
		return lib.mdPresortGetDestRateAuto3dig(self.I)

	def GetDestRateAutoADC(self):
		return lib.mdPresortGetDestRateAutoADC(self.I)

	def GetDestRateAutoMADC(self):
		return lib.mdPresortGetDestRateAutoMADC(self.I)

	def GetDestRate5dig(self):
		return lib.mdPresortGetDestRate5dig(self.I)

	def GetDestRate3dig(self):
		return lib.mdPresortGetDestRate3dig(self.I)

	def GetDestRateADC(self):
		return lib.mdPresortGetDestRateADC(self.I)

	def GetDestRateMADC(self):
		return lib.mdPresortGetDestRateMADC(self.I)

	def GetDestRateWS(self):
		return lib.mdPresortGetDestRateWS(self.I)

	def GetDestRateHDP(self):
		return lib.mdPresortGetDestRateHDP(self.I)

	def GetDestRateHD(self):
		return lib.mdPresortGetDestRateHD(self.I)

	def GetDestRateLOT(self):
		return lib.mdPresortGetDestRateLOT(self.I)

	def GetContainersTotal(self, p1):
		return lib.mdPresortGetContainersTotal(self.I, p1)

	def GetBundlesTotal(self, p1):
		return lib.mdPresortGetBundlesTotal(self.I, p1)

	def GetPiecesTotal(self, p1):
		return lib.mdPresortGetPiecesTotal(self.I, p1)

	def GetChargeTotal(self, p1):
		return lib.mdPresortGetChargeTotal(self.I, p1)

	def GetDestSCFInfo(self, p1):
		return lib.mdPresortGetDestSCFInfo(self.I, p1.encode('utf-8'))

	def GetDestNDCInfo(self, p1):
		return lib.mdPresortGetDestNDCInfo(self.I, p1.encode('utf-8'))

	def GetDestDDUInfo(self, p1):
		return lib.mdPresortGetDestDDUInfo(self.I, p1.encode('utf-8'))

	def GetDestContainersTotal(self, p1):
		return lib.mdPresortGetDestContainersTotal(self.I, p1)

	def GetDestBundlesTotal(self, p1):
		return lib.mdPresortGetDestBundlesTotal(self.I, p1)

	def GetDestPiecesTotal(self, p1):
		return lib.mdPresortGetDestPiecesTotal(self.I, p1)

	def GetDestChargeTotal(self, p1):
		return lib.mdPresortGetDestChargeTotal(self.I, p1)

	def SetTwoFtTrayMaximum(self, p1):
		lib.mdPresortSetTwoFtTrayMaximum(self.I, p1.encode('utf-8'))

	def SetTwoFtTrayMinimum(self, p1):
		lib.mdPresortSetTwoFtTrayMinimum(self.I, p1.encode('utf-8'))

	def SetOneFtTrayMaximum(self, p1):
		lib.mdPresortSetOneFtTrayMaximum(self.I, p1.encode('utf-8'))

	def SetOneFtTrayMinimum(self, p1):
		lib.mdPresortSetOneFtTrayMinimum(self.I, p1.encode('utf-8'))

	def SetBundleMaximum(self, p1):
		lib.mdPresortSetBundleMaximum(self.I, p1.encode('utf-8'))

	def SetPSSCFZipAutomationSort(self, p1):
		return lib.mdPresortSetPSSCFZipAutomationSort(self.I, p1.encode('utf-8'))

	def SetPSSCFZipNonAutomationSort(self, p1):
		return lib.mdPresortSetPSSCFZipNonAutomationSort(self.I, p1.encode('utf-8'))

	def SetPSSCFZipECRRTSort(self, p1):
		return lib.mdPresortSetPSSCFZipECRRTSort(self.I, p1.encode('utf-8'))

	def SetPSSCFZipCoSack(self, p1):
		return lib.mdPresortSetPSSCFZipCoSack(self.I, p1.encode('utf-8'))

	def SetPSNDCZipAutomationSort(self, p1):
		return lib.mdPresortSetPSNDCZipAutomationSort(self.I, p1.encode('utf-8'))

	def SetPSNDCZipNonAutomationSort(self, p1):
		return lib.mdPresortSetPSNDCZipNonAutomationSort(self.I, p1.encode('utf-8'))

	def SetPSNDCZipECRRTSort(self, p1):
		return lib.mdPresortSetPSNDCZipECRRTSort(self.I, p1.encode('utf-8'))

	def SetPSNDCZipCoSack(self, p1):
		return lib.mdPresortSetPSNDCZipCoSack(self.I, p1.encode('utf-8'))

	def SetPSDDUZipECRRTSort(self, p1):
		return lib.mdPresortSetPSDDUZipECRRTSort(self.I, p1.encode('utf-8'))

	def SetPSPOMZipAutomationSort(self, p1):
		return lib.mdPresortSetPSPOMZipAutomationSort(self.I, p1.encode('utf-8'))

	def SetPSPOMZipNonAutomationSort(self, p1):
		return lib.mdPresortSetPSPOMZipNonAutomationSort(self.I, p1.encode('utf-8'))

	def SetPSPOMZipECRRTSort(self, p1):
		return lib.mdPresortSetPSPOMZipECRRTSort(self.I, p1.encode('utf-8'))

	def SetPSPOMZipCoSack(self, p1):
		return lib.mdPresortSetPSPOMZipCoSack(self.I, p1.encode('utf-8'))

	def SetPSIncludeResidualPieces(self, p1):
		lib.mdPresortSetPSIncludeResidualPieces(self.I, p1)

	def ProducePostStatementForResidualPieces(self):
		return lib.mdPresortProducePostStatementForResidualPieces(self.I)

	def ProducePostStatement(self):
		lib.mdPresortProducePostStatement(self.I)

	def SetPostStatementToSelect(self, p1):
		lib.mdPresortSetPostStatementToSelect(self.I, p1)

	def SetIgnoreAspectRatio(self, p1):
		lib.mdPresortSetIgnoreAspectRatio(self.I, p1)

	def SetProduceIMBCode(self, p1):
		lib.mdPresortSetProduceIMBCode(self.I, p1)

	def GetIMBNumericCode(self):
		return lib.mdPresortGetIMBNumericCode(self.I).decode('utf-8')

	def GetIMBAlphaCode(self):
		return lib.mdPresortGetIMBAlphaCode(self.I).decode('utf-8')

	def GetIMBSerialNumber(self):
		return lib.mdPresortGetIMBSerialNumber(self.I).decode('utf-8')

	def SetPSMailingAgentCRID(self, p1):
		lib.mdPresortSetPSMailingAgentCRID(self.I, p1.encode('utf-8'))

	def SetContainerSequenceNumber(self, p1):
		return lib.mdPresortSetContainerSequenceNumber(self.I, p1.encode('utf-8'))

	def SetPSNonProfitAuthNumberMO(self, p1):
		lib.mdPresortSetPSNonProfitAuthNumberMO(self.I, p1.encode('utf-8'))

	def SetProduceDropShipForms(self, p1):
		return lib.mdPresortSetProduceDropShipForms(self.I, p1)

	def SetExpandedReportName(self, p1):
		lib.mdPresortSetExpandedReportName(self.I, p1)

	def ProduceMailDatFiles(self, p1, p2):
		return lib.mdPresortProduceMailDatFiles(self.I, p1.encode('utf-8'), p2.encode('utf-8'))

	def SetPSPostOfficeOfMailingPlus4(self, p1):
		lib.mdPresortSetPSPostOfficeOfMailingPlus4(self.I, p1.encode('utf-8'))

	def SetMDACSKeyLineData(self, p1):
		lib.mdPresortSetMDACSKeyLineData(self.I, p1.encode('utf-8'))

	def SetMDMachineID(self, p1):
		lib.mdPresortSetMDMachineID(self.I, p1.encode('utf-8'))

	def SetMDJobID(self, p1):
		lib.mdPresortSetMDJobID(self.I, p1.encode('utf-8'))

	def SetMDHDRIDEAllianceVersion(self, p1):
		lib.mdPresortSetMDHDRIDEAllianceVersion(self.I, p1.encode('utf-8'))

	def SetMDHDRHistorySequenceNumber(self, p1):
		lib.mdPresortSetMDHDRHistorySequenceNumber(self.I, p1.encode('utf-8'))

	def SetMDHDRHistoryStatus(self, p1):
		lib.mdPresortSetMDHDRHistoryStatus(self.I, p1.encode('utf-8'))

	def SetMDHDRHistoricalJobID(self, p1):
		lib.mdPresortSetMDHDRHistoricalJobID(self.I, p1.encode('utf-8'))

	def SetMDHDRLicensedUsersJobNumber(self, p1):
		lib.mdPresortSetMDHDRLicensedUsersJobNumber(self.I, p1.encode('utf-8'))

	def SetMDHDRJobNameTitleIssue(self, p1):
		lib.mdPresortSetMDHDRJobNameTitleIssue(self.I, p1.encode('utf-8'))

	def SetMDHDRFileSource(self, p1):
		lib.mdPresortSetMDHDRFileSource(self.I, p1.encode('utf-8'))

	def SetMDHDRUserLicenseCode(self, p1):
		lib.mdPresortSetMDHDRUserLicenseCode(self.I, p1.encode('utf-8'))

	def SetMDHDRMailDatSoftwareVendorName(self, p1):
		lib.mdPresortSetMDHDRMailDatSoftwareVendorName(self.I, p1.encode('utf-8'))

	def SetMDHDRMailDatSoftwareProductsName(self, p1):
		lib.mdPresortSetMDHDRMailDatSoftwareProductsName(self.I, p1.encode('utf-8'))

	def SetMDHDRMailDatSoftwareVersion(self, p1):
		lib.mdPresortSetMDHDRMailDatSoftwareVersion(self.I, p1.encode('utf-8'))

	def SetMDHDRMailDatSoftwareVendorEmail(self, p1):
		lib.mdPresortSetMDHDRMailDatSoftwareVendorEmail(self.I, p1.encode('utf-8'))

	def SetMDHDReDocSenderCRID(self, p1):
		lib.mdPresortSetMDHDReDocSenderCRID(self.I, p1.encode('utf-8'))

	def SetMDSEGVerificationFacilityName(self, p1):
		lib.mdPresortSetMDSEGVerificationFacilityName(self.I, p1.encode('utf-8'))

	def SetMDSEGVerificationFacilityZipPlus4(self, p1):
		lib.mdPresortSetMDSEGVerificationFacilityZipPlus4(self.I, p1.encode('utf-8'))

	def SetMDHDROriginalSoftwareVendorName(self, p1):
		lib.mdPresortSetMDHDROriginalSoftwareVendorName(self.I, p1.encode('utf-8'))

	def SetMDHDROriginalSoftwareProductsName(self, p1):
		lib.mdPresortSetMDHDROriginalSoftwareProductsName(self.I, p1.encode('utf-8'))

	def SetMDHDROriginalSoftwareVersion(self, p1):
		lib.mdPresortSetMDHDROriginalSoftwareVersion(self.I, p1.encode('utf-8'))

	def SetMDHDROriginalSoftwareVendorEmail(self, p1):
		lib.mdPresortSetMDHDROriginalSoftwareVendorEmail(self.I, p1.encode('utf-8'))

	def SetMDHDRContactName(self, p1):
		lib.mdPresortSetMDHDRContactName(self.I, p1.encode('utf-8'))

	def SetMDHDRContactPhone(self, p1):
		lib.mdPresortSetMDHDRContactPhone(self.I, p1.encode('utf-8'))

	def SetMDHDRContactEMail(self, p1):
		lib.mdPresortSetMDHDRContactEMail(self.I, p1.encode('utf-8'))

	def SetMDCPTMailOwnerID(self, p1):
		lib.mdPresortSetMDCPTMailOwnerID(self.I, p1.encode('utf-8'))

	def SetMDCPTOwnerCRID(self, p1):
		lib.mdPresortSetMDCPTOwnerCRID(self.I, p1.encode('utf-8'))

	def SetMDCPTMailOwnersMailingRefID(self, p1):
		lib.mdPresortSetMDCPTMailOwnersMailingRefID(self.I, p1.encode('utf-8'))

	def SetMDCPTPostalPriceIncID(self, p1):
		lib.mdPresortSetMDCPTPostalPriceIncID(self.I, p1.encode('utf-8'))

	def SetMDCPTPostalPriceIncType(self, p1):
		lib.mdPresortSetMDCPTPostalPriceIncType(self.I, p1.encode('utf-8'))

	def SetMDCPTStandParcelType(self, p1):
		lib.mdPresortSetMDCPTStandParcelType(self.I, p1.encode('utf-8'))

	def SetMDCPTStandFlatType(self, p1):
		lib.mdPresortSetMDCPTStandFlatType(self.I, p1.encode('utf-8'))

	def SetMDCPTUserOptField(self, p1):
		lib.mdPresortSetMDCPTUserOptField(self.I, p1.encode('utf-8'))

	def SetMDCPTContentOfMail(self, p1):
		lib.mdPresortSetMDCPTContentOfMail(self.I, p1.encode('utf-8'))

	def SetMDCSMCSAAgreementID(self, p1):
		lib.mdPresortSetMDCSMCSAAgreementID(self.I, p1.encode('utf-8'))

	def SetMDSEGDescription(self, p1):
		lib.mdPresortSetMDSEGDescription(self.I, p1.encode('utf-8'))

	def SetMDCPTComDescription(self, p1):
		lib.mdPresortSetMDCPTComDescription(self.I, p1.encode('utf-8'))

	def SetMDMPUName(self, p1):
		lib.mdPresortSetMDMPUName(self.I, p1.encode('utf-8'))

	def SetMDMPUDescription(self, p1):
		lib.mdPresortSetMDMPUDescription(self.I, p1.encode('utf-8'))

	def SetMDMPADescription(self, p1):
		lib.mdPresortSetMDMPADescription(self.I, p1.encode('utf-8'))

	def SetMDMPAMailingAgentMailerID(self, p1):
		lib.mdPresortSetMDMPAMailingAgentMailerID(self.I, p1.encode('utf-8'))

	def SetMDMPAMailOwnerPermitNumber(self, p1):
		lib.mdPresortSetMDMPAMailOwnerPermitNumber(self.I, p1.encode('utf-8'))

	def SetMDMPAMailOwnerPermitType(self, p1):
		lib.mdPresortSetMDMPAMailOwnerPermitType(self.I, p1.encode('utf-8'))

	def SetMDMPAAdditionalPostagePaymentMethod(self, val):
		lib.mdPresortSetMDMPAAdditionalPostagePaymentMethod(self.I, val.encode('utf-8'))

	def SetMDMPAAdditionalPaymentPermitNumber(self, val):
		lib.mdPresortSetMDMPAAdditionalPaymentPermitNumber(self.I, val.encode('utf-8'))

	def SetMDMPAAdditionalPaymentPermitZipPlus4(self, val):
		lib.mdPresortSetMDMPAAdditionalPaymentPermitZipPlus4(self.I, val.encode('utf-8'))

	def SetMDCCRCharacteristic(self, val):
		lib.mdPresortSetMDCCRCharacteristic(self.I, val.encode('utf-8'))

	def SetMDCCRCharacteristicType(self, val):
		lib.mdPresortSetMDCCRCharacteristicType(self.I, val.encode('utf-8'))

	def SetPOMLocaleKey(self, p1):
		lib.mdPresortSetPOMLocaleKey(self.I, p1.encode('utf-8'))

	def SetDDULocaleKey(self, p1):
		lib.mdPresortSetDDULocaleKey(self.I, p1.encode('utf-8'))

	def SetDDUPostalCode(self, p1):
		lib.mdPresortSetDDUPostalCode(self.I, p1.encode('utf-8'))
