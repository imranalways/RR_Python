from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait

import ConfigurationManager
from Connection import databaseConnection
from FindByXpath import by_xpath, select_dropdown_by_value, select_dropdown_text, click_xpath
from VisibilityCheck import supply_limit_visible


class Supply(object):

    def accept_alert_supply(self, driver, logging):
        try:
            WebDriverWait(driver, 5).until(ec.alert_is_present(), 'Waiting for alerts to appear in supplimentary')
            driver.switch_to.alert.accept()
            return 'accepted'
        except Exception as e:
            logging.warning('_From-> CardCreation()_ :' + str(e) + '\n supplementary alert accept error')
            return Exception

    def SupplyFields(self, results, driver, loan_id, connection_string, logging):

        driver.switch_to.alert.accept()
        limit_raise = False  # FLAG : for Credit limit exception throw

        for each_row in results[0:1]:

            try:
                s_date_field_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[2]/td[4]/input'
                # TODO supply submit date
                # s_date_field_value = time.strftime("%Y%m%d")
                processing_date = driver.find_element_by_xpath("/html/body/table[1]/tbody/tr/td[2]").text
                processing_date = processing_date.split(':', 1)
                processing_date = processing_date[1].split('\n', 1)
                processing_date = processing_date[0].strip()

                s_date_field_value = processing_date
                by_xpath(s_date_field_xpath, s_date_field_value, driver)

                s_fee_code_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[6]/td[2]/select'
                s_fee_code_value = each_row['BSFeeCode']
                select_dropdown_by_value(s_fee_code_xpath, s_fee_code_value, driver)

                s_credit_limit_name = '//*[@id="varSt_request_credit_limit1"]'
                s_credit_limit = each_row['PRDCreditLimit']
                by_xpath(s_credit_limit_name, s_credit_limit, driver)

                # region PersonalReferance

                s_relationship = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[9]/td[2]/select'
                s_relationship_value = each_row['PRDRelationship']
                select_dropdown_by_value(s_relationship, s_relationship_value, driver)

                s_title = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[10]/td[2]/select'
                s_title_value = each_row['PRDTitle']
                select_dropdown_by_value(s_title, s_title_value, driver)

                s_fullname = '//*[@id="varSt_card_fullname"]'
                s_fullname_value = each_row['PRDFullName']
                by_xpath(s_fullname, s_fullname_value, driver)

                s_aliastitle = '//*[@id="varSt_alias_title"]'
                s_aliastitle_value = each_row['PRDAliasTitle']
                by_xpath(s_aliastitle, s_aliastitle_value, driver)

                s_aliasname = '//*[@id="varSt_alias_name"]'
                s_aliasname_value = each_row['PRDAliasName']
                by_xpath(s_aliasname, s_aliasname_value, driver)

                s_p_address1 = '//*[@id="varSt_home_addr1"]'
                s_p_address1_value = each_row['PRDAddress1']
                by_xpath(s_p_address1, s_p_address1_value, driver)

                s_p_address2 = '//*[@id="varSt_home_addr2"]'
                s_p_address2_value = each_row['PRDAddress2']
                by_xpath(s_p_address2, s_p_address2_value, driver)

                s_p_address3 = '//*[@id="varSt_home_addr3"]'
                s_p_address3_value = each_row['PRDAddress3']
                by_xpath(s_p_address3, s_p_address3_value, driver)

                s_p_address4 = '//*[@id="varSt_home_addr4"]'
                s_p_address4_value = each_row['PRDAddress4']
                by_xpath(s_p_address4, s_p_address4_value, driver)

                s_p_address5 = '//*[@id="varSt_home_addr5"]'
                s_p_address5_value = each_row['PRDAddress5']
                by_xpath(s_p_address5, s_p_address5_value, driver)

                s_emboss_name = '//*[@id="varSt_embossname"]'
                s_emboss_name_value = each_row['PRDEmbossName']
                by_xpath(s_emboss_name, s_emboss_name_value, driver)

                s_emboss_name2 = '//*[@id="varSt_embossname2"]'
                s_emboss_name2_value = each_row['PRDEmbossName2']
                by_xpath(s_emboss_name2, s_emboss_name2_value, driver)

                s_old_nric = '//*[@id="varSt_old_ic_no"]'
                s_old_nric_value = each_row['loanId']
                by_xpath(s_old_nric, s_old_nric_value, driver)

                s_cif_no = '//*[@id="varSt_cif_no"]'
                s_cif_no_value = each_row['PRDCIFNo']
                by_xpath(s_cif_no, s_cif_no_value, driver)

                s_pp_no = '//*[@id="varSt_passport_no"]'
                s_pp_value = each_row['PRDPassportNumber']
                by_xpath(s_pp_no, s_pp_value, driver)

                s_gender = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[17]/td[2]/select'
                s_gender_value = each_row['PRDGender']
                select_dropdown_text(s_gender, s_gender_value, driver)

                s_ethnic = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[17]/td[4]/select'
                s_ethnic_value = each_row['PRDEthnic']
                select_dropdown_by_value(s_ethnic, s_ethnic_value, driver)

                s_dob = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[18]/td[2]/input'
                s_dob_value = each_row['PRDDateOfBirth']
                by_xpath(s_dob, s_dob_value, driver)

                s_telephone = '// *[ @ id = "varSt_home_telno"]'
                s_telephone_value = each_row['PRDTelephone']
                by_xpath(s_telephone, s_telephone_value, driver)

                s_mobile = '//*[@id="varSt_mobile_no"]'
                s_mobile_value = each_row['PRDMobile']
                by_xpath(s_mobile, s_mobile_value, driver)

                s_pager = '//*[@id="varSt_pager_no"]'
                s_pager_value = each_row['PRDPager']
                by_xpath(s_pager, s_pager_value, driver)

                s_email = '//*[@id="varSt_email"]'
                s_email_value = each_row['PRDEmail']
                by_xpath(s_email, s_email_value, driver)

                s_country = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[18]/td[4]/select'
                s_country_value = each_row['PRDCountry']
                select_dropdown_by_value(s_country, s_country_value, driver)

                s_state = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[19]/td[4]/select'
                s_state_value = each_row['PRDState']
                select_dropdown_by_value(s_state, s_state_value, driver)

                if each_row['PRDState'] != each_row['PRDCity']:
                    pass
                else:
                    s_city = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[20]/td[4]/select'
                    s_city_value = each_row['PRDCity']
                    select_dropdown_by_value(s_city, s_city_value, driver)

                s_postcode = '//*[@id="varSt_home_postcode"]'
                s_postcode_value = each_row['PRDPostCode']
                by_xpath(s_postcode, s_postcode_value, driver)

                s_mothermaiden_name = '//*[@id="varSt_mother_maiden_name"]'
                s_mothermaiden_name_value = each_row['PRDMotherMaidenName']
                by_xpath(s_mothermaiden_name, s_mothermaiden_name_value, driver)

                s_residentcountry = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[23]/td[2]/select'
                s_residentcountry_value = each_row['PRDResidentCountry']
                select_dropdown_by_value(s_residentcountry, s_residentcountry_value, driver)

                s_prstatus = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[24]/td[2]/select'
                s_s_prstatus_value = each_row['PRDPRStatus']
                select_dropdown_by_value(s_prstatus, s_s_prstatus_value, driver)

                # TODO :resident period need to be checked
                # s_Residentperiod1 = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[25]/td[2]/input[1]'
                # s_Residentperiod1Value = each_row['PRDResidentPeriod']
                # PathObject .by_xpath(s_Residentperiod1, s_Residentperiod1Value, driver)
                # s_Residentperiod2 = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[25]/td[2]/input[2]'
                # s_Residentperiod2Value = each_row['PRDResidentPeriod']
                # PathObject .by_xpath(s_Residentperiod2, s_Residentperiod2Value, driver)

                s_social_status = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[26]/td[2]/select'
                s_social_status_value = each_row['PRDSocialStatus']
                select_dropdown_by_value(s_social_status, s_social_status_value, driver)

                s_homeownership = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[27]/td[2]/select'
                s_homeownership_value = each_row['PRDHomeOwnership']
                select_dropdown_by_value(s_homeownership, s_homeownership_value, driver)

                s_housetype = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[28]/td[2]/select'
                s_housetype_value = each_row['PRDHouseType']
                select_dropdown_by_value(s_housetype, s_housetype_value, driver)

                s_maritalstatus = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[23]/td[4]/select'
                s_maritalstatus_value = each_row['PRDMaritalStatus']
                select_dropdown_text(s_maritalstatus, s_maritalstatus_value, driver)

                s_qualification = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[24]/td[4]/select'
                s_qualification_value = each_row['PRDQualification']
                select_dropdown_by_value(s_qualification, s_qualification_value, driver)

                s_microfilm = '//*[@id="varSt_microfilm_no"]'
                s_microfilm_value = each_row['PRDMicrofilm']
                by_xpath(s_microfilm, s_microfilm_value, driver)

                # s_nationality = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[26]/td[4]/select'
                s_nationality = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[26]/td[4]/select'
                s_nationalityValue = each_row['PRDNationality']
                select_dropdown_by_value(s_nationality, s_nationalityValue, driver)

                s_liquidAsset = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[27]/td[4]/select'
                s_liquidAssetValue = each_row['PRDLiquidAsset']
                select_dropdown_by_value(s_liquidAsset, s_liquidAssetValue, driver)

                # endregion

                # region AlternativeAddress1

                s_Address1 = '//*[@id="varSt_alt1_bill_addr1"]'
                s_Address1Value = each_row['A1AAddress1']
                by_xpath(s_Address1, s_Address1Value, driver)

                s_Address2 = '//*[@id="varSt_alt1_bill_addr2"]'
                s_Address2Value = each_row['A1AAddress2']
                by_xpath(s_Address2, s_Address2Value, driver)

                s_Address3 = '//*[@id="varSt_alt1_bill_addr3"]'
                s_Address3Value = each_row['A1AAddress3']
                by_xpath(s_Address3, s_Address3Value, driver)

                s_Address4 = '//*[@id="varSt_alt1_bill_addr4"]'
                s_Address4Value = each_row['A1AAddress4']
                by_xpath(s_Address4, s_Address4Value, driver)

                s_Address5 = '//*[@id="varSt_alt1_bill_addr5"]'
                s_Address5Value = each_row['A1AAddress5']
                by_xpath(s_Address5, s_Address5Value, driver)

                s_country = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[35]/td[2]/select'
                s_country_value = each_row['A1ACountry']
                select_dropdown_by_value(s_country, s_country_value, driver)

                s_state = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[36]/td[2]/select'
                s_state_value = each_row['A1AState']
                select_dropdown_by_value(s_state, s_state_value, driver)

                if each_row['A1AState'] != each_row['A1ACity']:
                    pass
                else:
                    s_city = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[37]/td[2]/select'
                    s_city_value = each_row['A1ACity']
                    select_dropdown_by_value(s_city, s_city_value, driver)

                s_postcode = '//*[@id="varSt_alt1_bill_postcode"]'
                s_postcode_value = each_row['A1APostCode']
                by_xpath(s_postcode, s_postcode_value, driver)

                # endregion

                # region AlternativeAddress2

                s_Address1 = '//*[@id="varSt_alt2_bill_addr1"]'
                s_Address1Value = each_row['A2AAddress1']
                by_xpath(s_Address1, s_Address1Value, driver)

                s_Address2 = '//*[@id="varSt_alt2_bill_addr2"]'
                s_Address2Value = each_row['A2AAddress2']
                by_xpath(s_Address2, s_Address2Value, driver)

                s_Address3 = '//*[@id="varSt_alt2_bill_addr3"]'
                s_Address3Value = each_row['A2AAddress3']
                by_xpath(s_Address3, s_Address3Value, driver)

                s_Address4 = '//*[@id="varSt_alt2_bill_addr4"]'
                s_Address4Value = each_row['A2AAddress4']
                by_xpath(s_Address4, s_Address4Value, driver)

                s_Address5 = '//*[@id="varSt_alt2_bill_addr5"]'
                s_Address5Value = each_row['A2AAddress5']
                by_xpath(s_Address5, s_Address5Value, driver)

                s_country = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[35]/td[4]/select'
                s_country_value = each_row['A2ACountry']
                select_dropdown_by_value(s_country, s_country_value, driver)

                s_state = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[36]/td[4]/select'
                s_state_value = each_row['A2AState']
                select_dropdown_by_value(s_state, s_state_value, driver)

                if each_row['A2AState'] != each_row['A2ACity']:
                    pass
                else:
                    s_city = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[37]/td[4]/select'
                    s_city_value = each_row['A2ACity']
                    select_dropdown_by_value(s_city, s_city_value, driver)

                s_postcode = '//*[@id="varSt_alt2_bill_postcode"]'
                s_postcode_value = each_row['A2APostCode']
                by_xpath(s_postcode, s_postcode_value, driver)

                # endregion

                # region CompanyReferance

                s_company = '//*[@id="varSt_comp_name"]'
                s_companyValue = each_row['CRCompanyName']
                by_xpath(s_company, s_companyValue, driver)

                s_Address1 = '//*[@id="varSt_comp_addr1"]'
                s_Address1Value = each_row['CRAddress1']
                by_xpath(s_Address1, s_Address1Value, driver)

                s_Address2 = '//*[@id="varSt_comp_addr2"]'
                s_Address2Value = each_row['CRAddress2']
                by_xpath(s_Address2, s_Address2Value, driver)

                s_Address3 = '//*[@id="varSt_comp_addr3"]'
                s_Address3Value = each_row['CRAddress3']
                by_xpath(s_Address3, s_Address3Value, driver)

                s_Address4 = '//*[@id="varSt_comp_addr4"]'
                s_Address4Value = each_row['CRAddress4']
                by_xpath(s_Address4, s_Address4Value, driver)

                s_Address5 = '//*[@id="varSt_comp_addr5"]'
                s_Address5Value = each_row['CRAddress5']
                by_xpath(s_Address5, s_Address5Value, driver)

                s_country = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[46]/td[2]/select'
                s_country_value = each_row['CRCountry']
                select_dropdown_by_value(s_country, s_country_value, driver)

                s_state = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[47]/td[2]/select'
                s_state_value = each_row['CRState']
                select_dropdown_by_value(s_state, s_state_value, driver)

                if each_row['CRState'] != each_row['CRCity']:
                    pass
                else:
                    s_city = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[48]/td[2]/select'
                    s_city_value = each_row['CRCity']
                    select_dropdown_by_value(s_city, s_city_value, driver)

                s_postcode = '//*[@id="varSt_comp_postcode"]'
                s_postcode_value = each_row['CRPostCode']
                by_xpath(s_postcode, s_postcode_value, driver)

                s_telephone = '//*[@id="varSt_comp_telno"]'
                s_telephone_value = each_row['CRTelephone']
                by_xpath(s_telephone, s_telephone_value, driver)

                s_organizationType = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[51]/td[2]/select'
                s_organizationTypeValue = each_row['CROrganisationType']
                select_dropdown_by_value(s_organizationType, s_organizationTypeValue, driver)

                s_creditstatus = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[52]/td[2]/select'
                s_creditstatusValue = each_row['OLCreditStatus']
                select_dropdown_by_value(s_creditstatus, s_creditstatusValue, driver)

                s_fax = '//*[@id="varSt_comp_faxno"]'
                s_faxValue = each_row['CRFax']
                by_xpath(s_fax, s_faxValue, driver)

                s_designation = '//*[@id="varSt_comp_desgn"]'
                s_designationValue = each_row['CRDesignation']
                by_xpath(s_designation, s_designationValue, driver)

                s_Occupation = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[55]/td[2]/select'
                s_OccupationValue = each_row['CROccupation']
                select_dropdown_by_value(s_Occupation, s_OccupationValue, driver)

                # endregion

                # region ApplicationReferance

                s_billingcode = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[40]/td[4]/select'
                s_billingcodeValue = each_row['PCRBillingCode']
                select_dropdown_by_value(s_billingcode, s_billingcodeValue, driver)

                s_seperatedStatement = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[41]/td[4]/select'
                s_seperatedStatementValue = each_row['PCRSeparatedStatement']
                select_dropdown_by_value(s_seperatedStatement, s_seperatedStatementValue, driver)

                s_cardCollectionMethod = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[42]/td[4]/select'
                s_cardCollectionMethodValue = each_row['PCRCardCollectionMethod']
                select_dropdown_by_value(s_cardCollectionMethod, s_cardCollectionMethodValue, driver)

                s_cardDeliveryAddressCode = '//*[@id="idDelvery2"]/select'
                s_cardDeliveryAddressCodeValue = each_row['PCRCardDeliveryAddressCode']
                select_dropdown_by_value(s_cardDeliveryAddressCode, s_cardDeliveryAddressCodeValue, driver)

                s_legalBillingAddressCode = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[44]/td[4]/select'
                s_legalBillingAddressCodeValue = each_row['PCRLegalBillingAddressCode']
                select_dropdown_by_value(s_legalBillingAddressCode, s_legalBillingAddressCodeValue, driver)

                s_customerClass = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[45]/td[4]/select'
                s_customerClassValue = each_row['OCustomerClass']
                select_dropdown_by_value(s_customerClass, s_customerClassValue, driver)

                s_interestGroup = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[46]/td[4]/select'
                s_interestGroupValue = each_row['OInterestGroup']
                select_dropdown_by_value(s_interestGroup, s_interestGroupValue, driver)

                s_documentComplete = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[47]/td[4]/select'
                s_documentCompleteValue = each_row['ODocumentComplete']
                select_dropdown_by_value(s_documentComplete, s_documentCompleteValue, driver)

                s_photoIndicator = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[48]/td[4]/select'
                s_photoIndicatorValue = each_row['OPhotoIndicator']
                select_dropdown_by_value(s_photoIndicator, s_photoIndicatorValue, driver)

                s_languageIndicator = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[49]/td[4]/select'
                s_languageIndicatorValue = each_row['OLanguageIndicator']
                select_dropdown_by_value(s_languageIndicator, s_languageIndicatorValue, driver)

                s_pinGenIndicator = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[50]/td[4]/select'
                s_pinGenIndicatorValue = each_row['OATMIndicator']
                select_dropdown_by_value(s_pinGenIndicator, s_pinGenIndicatorValue, driver)

                s_jobStability = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[51]/td[4]/select'
                s_jobStabilityValue = each_row['CRJobStabiliy']
                select_dropdown_by_value(s_jobStability, s_jobStabilityValue, driver)

                # TODO future update reward acount
                # s_rewardPoint = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[2]/td[4]/input'
                # s_rewardPointValue = each_row['CRJobStabiliy']
                # PathObject .by_xpath(s_rewardPoint, s_rewardPointValue, driver)

                # endregion

                # region AdditionalInformation

                S_KycDate = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[94]/td[4]/input'
                S_KycDateValue = each_row['AIDKycNextReviewDate']
                by_xpath(S_KycDate, S_KycDateValue, driver)

                S_KycRisk = '//*[@id="varCb_adl_field_num01"]'
                S_KycRiskValue = each_row['AINKycRiskInd']
                by_xpath(S_KycRisk, S_KycRiskValue, driver)

                s_path_object = '//*[@id="varCb_adl_field_name01"]'
                s_xpath_value = each_row['father']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name02"]'
                s_xpath_value = each_row['NID']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name03"]'
                s_xpath_value = each_row['AICApplicationSerialNo']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name04"]'
                s_xpath_value = each_row['AICSectorCode']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name05"]'
                s_xpath_value = each_row['TinNo']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name06"]'
                s_xpath_value = each_row['AICLienAmtAgainstSecuredCard']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name07"]'
                s_xpath_value = each_row['AICPriorityPass']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name08"]'
                s_xpath_value = each_row['AICPassportDetail']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name09"]'
                s_xpath_value = each_row['AICPassportIssued']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name10"]'
                s_xpath_value = each_row['AICPassportExpiry']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name11"]'
                s_xpath_value = each_row['AICPassportIssuePlace']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name13"]'
                s_xpath_value = each_row['AICSuppleRecommenderNumber']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name14"]'
                s_xpath_value = each_row['AICCompanyType']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name17"]'
                s_xpath_value = each_row['AICApprovalLevel']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name18"]'
                s_xpath_value = each_row['AICApproverPIN']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name19"]'
                s_xpath_value = each_row['AICAnalystPIN']
                by_xpath(s_path_object, s_xpath_value, driver)

                s_path_object = '//*[@id="varCb_adl_field_name20"]'
                s_xpath_value = each_row['AICSubSectorCodeSSS']
                by_xpath(s_path_object, s_xpath_value, driver)

                # endregion

                S_AddXpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[2]/tbody/tr/td/input[1]'
                click_xpath(S_AddXpath, driver)

                driver.switch_to.alert.accept()

                S_PathObject = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[4]/td[2]/select'
                S_XpathValue = "1-Approved"
                select_dropdown_text(S_PathObject, S_XpathValue, driver)

                try:
                    s_local_limit_xpath = '//*[@id="varSt_local_credit_limit"]'
                    s_foreign_credit_xpath = '//*[@id="varSt_foreign_credit_limit"]'
                    s_local_limit_name = 'varSt_local_credit_limit'
                    s_foreign_credit_name = 'varSt_foreign_credit_limit'

                    s_credit_limit = each_row['PRDCreditLimit']

                    try:
                        lc_visible = driver.find_element_by_name(s_local_limit_name).is_displayed()
                    except:
                        lc_visible = False

                    try:
                        fc_visible = driver.find_element_by_name(s_foreign_credit_name).is_displayed()
                    except:
                        fc_visible = False

                    if lc_visible is True and fc_visible is True:
                        by_xpath(s_local_limit_xpath, s_credit_limit, driver)
                    elif lc_visible is True:
                        by_xpath(s_local_limit_xpath, s_credit_limit, driver)
                    elif fc_visible is True:
                        by_xpath(s_foreign_credit_xpath, s_credit_limit, driver)

                    # MonitorCode
                    S_PathObject = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[12]/td[2]/select'
                    S_XpathValue = each_row['NPMonitorCode']
                    select_dropdown_by_value(S_PathObject, S_XpathValue, driver)

                    loanIdSupply = each_row['loanId']
                    application_no = driver.find_element_by_xpath(
                        "/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[2]/td[2]").text

                    S_AddXpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[2]/tbody/tr/td/input[1]'
                    click_xpath(S_AddXpath, driver)

                    driver.switch_to.alert.accept()
                    # time.sleep(1)

                    isTrue = False
                    while isTrue is False:
                        isTrue = supply_limit_visible(driver, logging)

                        if isTrue is True:
                            raise Exception
                        else:
                            break

                except Exception as e:
                    logging.warning('_From-> SupplementaryCard()_ :' + str(e) + '\n supplementary credit page error')
                    loan_id = each_row['loanId']
                    procedureName = "EXEC RDMS_UpdateCardCreditError_Card_Rpa @loanId=?"
                    objectConnection = databaseConnection()
                    objectConnection.callProcedureError(procedureName, connection_string, loan_id)
                    results = []

                    SupplyErrorURL = ConfigurationManager.RobotData("SupplyErrorURL")

                    driver.get(SupplyErrorURL)
                    limit_raise = True
                    raise Exception

                procedureName = "EXEC RDMS_UpdateApplicationNo_Card_Rpa @loanId=?, @ApplicationNo=?"
                objectConnection = databaseConnection()
                objectConnection.callUpdateProcedure(procedureName, connection_string, loanIdSupply, application_no)

            except Exception as e:
                logging.warning('_From-> SupplementaryCard()_ :' + str(e) + '\n supplementary field error')
                if limit_raise:
                    limit_raise = False
                    raise Exception
                else:
                    Accpted = self.accept_alert_supply(driver, logging)
                    loanIdSupply = each_row['loanId']
                    procedureName = "EXEC RDMS_UpdateCardNoError_Card_Rpa @loanId=?"
                    objectConnection = databaseConnection()
                    objectConnection.callProcedureError(procedureName, connection_string, loanIdSupply)
                    driver.refresh()
                    raise Exception

        results = []

        return results

