from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select

from Connection import databaseConnection
from SupplementaryCard import Supply
from VisibilityCheck import basic_limit_visible


def basic_card(driver, results, connection_string, logging):
    limit_raise = False  # FLAG : for Credit limit exception throw
    add1_button_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[4]/tbody/tr/td/input[1]'
    WebDriverWait(driver, 1).until(lambda chrome_driver: driver.find_element_by_xpath(add1_button_xpath)).click()
    driver.switch_to.alert.accept()

    for each_row in results[0:1]:  # loop for basic card

        try:

            # TODO idtype nric
            # id_type_nric = '123456'
            # id_type_nric_name ='//*[@id="varBt_nric_no"]'
            #
            # FindByXpath.byXpath(id_type_nric_name, id_type_nric, driver)
            # Idtype_nricElement = WebDriverWait(driver, 1).until(lambda chrome_driver:
            #                                                     driver.find_element_by_xpath(id_type_nric_name))
            # Idtype_nricElement.send_keys(id_type_nric)

            select = Select(driver.find_element_by_name('varBt_ser_centre_code'))
            select.select_by_value(each_row['BSBranchServiceCenter'])

            select = Select(driver.find_element_by_name('varBt_card_prodct_group'))
            select.select_by_value(each_row['BSProductGroup'])

            # TODO :mark program
            # mark_program = '123456'
            # mark_program_name = 'varBt_mkt_program_code'
            # mark_program_element = WebDriverWait(driver, 1).\
            #     until(lambda driver: driver.find_element_by_name(mark_program_name))
            # mark_program_element.send_keys(mark_program)
            # TODO: AUTO generate submit date

            processing_date = driver.find_element_by_xpath("/html/body/table[1]/tbody/tr/td[2]").text
            processing_date = processing_date.split(':', 1)
            processing_date = processing_date[1].split('\n', 1)
            processing_date = processing_date[0].strip()

            # submitted_date = time.strftime("%Y%m%d")
            # '20010921'

            date_button_name = 'varBt_submitted_date'
            date_button_element = WebDriverWait(driver, 1).until(
                lambda chrome_driver: driver.find_element_by_name(date_button_name))
            date_button_element.send_keys(processing_date)

            # TODO: add corporate id in database
            # corporate_id = '123'
            # corporate_id_name = 'varBt_corp_id'
            # corporate_idElement = WebDriverWait(driver, 1).\
            #     until(lambda chrome_driver: driver.find_element_by_name(corporate_id_name))
            # corporate_idElement.send_keys(corporate_id)

            fee_code = each_row['BSFeeCode']
            if fee_code is not None and fee_code.strip() != '':
                select = Select(driver.find_element_by_name('var_feeCode'))
                select.select_by_value(fee_code)
            else:
                pass

            # TODO: add application source code
            # click_SearchXpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[5]/td[4]/input[2]'
            # WebDriverWait(driver, 1).until(lambda driver: driver.find_element_by_xpath(click_SearchXpath)).click()
            # TODO: add gift code
            # select = Select(driver.find_element_by_name('varBt_gift_code'))
            # select.select_by_visible_text(each_row['varBt_gift_code'])

            # region PERSONAL DETAILS

            if each_row['PRDTitle'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('var_titleCode'))
                select.select_by_value(each_row['PRDTitle'])

            fullname = each_row['PRDFullName']
            full_button_name = 'varBt_card_fullname'
            full_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(full_button_name))
            full_button_element.send_keys(fullname)

            if each_row['PRDAliasTitle'] is None:
                pass
            elif each_row['PRDAliasTitle'].strip() == '' or each_row['PRDAliasTitle'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_alias_title'))
                select.select_by_value(each_row['PRDAliasTitle'])

            alias_button_name = 'varBt_alias_name'
            alias_name = each_row['PRDAliasName']
            alias_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(alias_button_name))
            alias_button_element.send_keys(alias_name)

            emboss_name = each_row['PRDEmbossName']
            emboss_button_name = 'varBt_embossname'
            emboss_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(emboss_button_name))
            emboss_button_element.send_keys(emboss_name)

            emboss_2nd_button_name = 'varBt_embossname2'
            emboss_name_2 = each_row['PRDEmbossName2']
            if emboss_name_2 is not None and emboss_name_2.strip() != '':
                emboss_2nd_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(emboss_2nd_button_name))
                emboss_2nd_button_element.send_keys(emboss_name_2)
            else:
                pass

            nric_no = each_row['loanID']
            nric_button_name = 'varBt_old_ic_no'
            if nric_button_name is not None and nric_button_name.strip() != '':
                nric_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(nric_button_name))
                nric_button_element.send_keys(nric_no)
            else:
                pass

            cif_no = each_row['PRDCIFNo']
            cif_button_name = 'varBt_cif_no'
            cif_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(cif_button_name))
            cif_button_element.send_keys(cif_no)

            passport_no = each_row['PRDPassportNumber']
            passport_no_button_name = 'varBt_passport_no'
            if passport_no is not None and passport_no.strip() != '':
                passport_no_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(passport_no_button_name))
                passport_no_button_element.send_keys(passport_no)
            else:
                pass

            nationality = each_row['PRDNationality']
            nationality_button_name = 'varBt_nationality'
            nationality_button_element = Select(driver.find_element_by_name(nationality_button_name))
            nationality_button_element.select_by_value(nationality)

            # NOTE:
            # connect plus & card pro dropdown value code difference
            # SO MATCHED BY TEXT
            gender = each_row['PRDGender']
            if gender == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_sex'))
                select.select_by_visible_text(gender)

            dob = each_row['PRDDateOfBirth']
            dob_button_name = 'varBt_dob'
            dob_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(dob_button_name))
            dob_button_element.send_keys(dob)

            telephone_no = each_row['PRDTelephone']
            telephone_button_name = 'varBt_home_telno'
            if telephone_no is not None and telephone_no.strip() != '':
                telephone_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(telephone_button_name))
                telephone_button_element.send_keys(telephone_no)
            else:
                pass

            mobile_button_name = 'varBt_mobile_no'
            mobile_no = each_row['PRDMobile']
            if mobile_no is not None and mobile_no.strip() != '':
                mobile_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(mobile_button_name))
                mobile_button_element.send_keys(mobile_no)
            else:
                pass

            pager_button_name = 'varBt_pager_no'
            pager = each_row['PRDPager']
            if pager is not None and pager.strip() != '':
                pager_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(pager_button_name))
                pager_button_element.send_keys(pager)
            else:
                pass

            address_1 = each_row['PRDAddress1']
            address_1st_button_name = 'varBt_home_addr1'
            address_1st_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(address_1st_button_name))
            address_1st_button_element.send_keys(address_1)

            address_2 = each_row['PRDAddress2']
            address_2nd_button_name = 'varBt_home_addr2'
            if address_2 is not None and address_2.strip() != '':
                address_2nd_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(address_2nd_button_name))
                address_2nd_button_element.send_keys(address_2)
            else:
                pass

            address_3 = each_row['PRDAddress3']
            address_3rd_button_name = 'varBt_home_addr3'
            if address_3 is None or address_3.strip() == '':
                pass
            else:
                address_3rd_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(address_3rd_button_name))
                address_3rd_button_element.send_keys(address_3)

            address_4 = each_row['PRDAddress4'].strip()
            address_4th_button_name = 'varBt_home_addr4'
            if address_4 is not None and address_4 != '':
                address_4th_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(address_4th_button_name))
                address_4th_button_element.send_keys(address_4)
            else:
                pass

            address_5 = each_row['PRDAddress5']
            address_5th_button_name = 'varBt_home_addr5'
            if address_5 is not None and address_5.strip() != '':
                address_5th_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(address_5th_button_name))
                address_5th_button_element.send_keys(address_5)
            else:
                pass

            ethnic = each_row['PRDEthnic']
            if ethnic is None or ethnic.strip() == '' or ethnic == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_ethnic'))
                select.select_by_value(ethnic)

            if each_row['PRDCountry'] is not None and each_row['PRDCountry'].strip() != '' \
                    and each_row['PRDCountry'] != 'NS':
                select = Select(driver.find_element_by_name('varBt_home_cntry_cd'))
                select.select_by_value(each_row['PRDCountry'])
            else:
                pass

            if each_row['PRDState'] is None or each_row['PRDState'].strip() == '' \
                    or each_row['PRDState'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_home_state'))
                select.select_by_value(each_row['PRDState'])

            if each_row['PRDCity'] is None or each_row['PRDCity'].strip() == '' or \
                    each_row['PRDState'] != each_row['PRDCity'] or each_row['PRDCity'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_home_city'))
                select.select_by_value(each_row['PRDCity'])

            postcode = each_row['PRDPostCode']
            postcode_button_name = 'varBt_home_postcode'
            if postcode is None or postcode.strip() == '':
                pass
            else:
                postcode_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(postcode_button_name))
                postcode_button_element.send_keys(postcode)

            email_button_name = 'varBt_email'
            email = each_row['PRDEmail']
            email_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(email_button_name))
            email_button_element.send_keys(email)

            email2 = each_row['PRDEmail2']
            email_button_name2 = 'varBt_email2'
            if email2 is None or email2.strip() == '':
                pass
            else:
                email_button2_element = WebDriverWait(driver, 1).until(lambda chromedriver:
                                                                       driver.find_element_by_name(email_button_name2))
                email_button2_element.send_keys(email2)

            email3 = each_row['PRDEmail3']
            email_button_name3 = 'varBt_email3'
            if email3 is None or email3.strip() == '':
                pass
            else:
                email_button3_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(email_button_name3))
                email_button3_element.send_keys(email3)

            if each_row['PRDResidentCountry'] is None or each_row['PRDResidentCountry'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_residency_code'))
                select.select_by_value(each_row['PRDResidentCountry'])

            if each_row['PRDMaritalStatus'] is None or each_row['PRDMaritalStatus'].strip() == '' \
                    or each_row['PRDMaritalStatus'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_marital_status'))
                select.select_by_visible_text(each_row['PRDMaritalStatus'])

            if each_row['PRDHomeOwnership'] is None or each_row['PRDHomeOwnership'].strip() == '' \
                    or each_row['PRDHomeOwnership'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_home_ownership'))
                select.select_by_value(each_row['PRDHomeOwnership'])

            if each_row['PRDHouseType'] is None or each_row['PRDHouseType'].strip() == '' \
                    or each_row['PRDHouseType'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_house_type'))
                select.select_by_value(each_row['PRDHouseType'])

            if each_row['PRDLiquidAsset'] is None or each_row['PRDLiquidAsset'].strip() == '' \
                    or each_row['PRDLiquidAsset'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_liquid_asset'))
                select.select_by_value(each_row['PRDLiquidAsset'])

            if each_row['PRDSocialStatus'] is None or each_row['PRDSocialStatus'].strip() == '' \
                    or each_row['PRDSocialStatus'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_social_status'))
                select.select_by_value(each_row['PRDSocialStatus'])

            if each_row['PRDPRStatus'] is None or each_row['PRDPRStatus'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_pr_of_country'))
                select.select_by_value(each_row['PRDPRStatus'])

            no_of_dependents = each_row['PRDNoOfDependents'].strip()
            dependents_button_name = 'varBt_no_of_dependents'
            if no_of_dependents is None or no_of_dependents == '':
                pass
            else:
                dependents_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(dependents_button_name))
                dependents_button_element.send_keys(no_of_dependents)

            resident_button_name = 'varBt_year_there_year'
            resident_period = each_row['PRDResidentPeriod']
            if resident_period is None or resident_period.strip() == '':
                pass
            else:
                resident_period = resident_period.split(' ', 1)
                resident_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(resident_button_name))
                driver.find_element_by_name(resident_button_name).clear()
                resident_button_element.send_keys(resident_period[0])

            resident2_button_name = 'varBt_year_there_month'
            resident_period2 = each_row['PRDResidentPeriod']
            if resident_period2 is None or resident_period2.strip() == '':
                pass
            else:
                resident_period2 = resident_period2.split(' ', 1)
                resident_button2_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(resident2_button_name))
                driver.find_element_by_name(resident2_button_name).clear()
                resident_button2_element.send_keys(resident_period2[1])

            if each_row['PRDQualification'] is None or each_row['PRDQualification'].strip() == '' \
                    or each_row['PRDQualification'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_qualification'))
                select.select_by_value(each_row['PRDQualification'])

            microfilm = each_row['PRDMicrofilm']
            microfilm_button_name = 'varBt_microfilm_no'
            if microfilm is None or microfilm.strip() == '':
                pass
            else:
                microfilm_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(microfilm_button_name))
                microfilm_button_element.send_keys(microfilm)

            mother_name = each_row['PRDMotherMaidenName']
            mother_button_name = 'varBt_mother_maiden_name'
            mother_button_element = WebDriverWait(driver, 1).until(
                lambda chromedriver: driver.find_element_by_name(mother_button_name))
            mother_button_element.send_keys(mother_name)
            # endregion

            # region ALTERNATIVE-1 ADDRESS

            alt1_address_1st_button_name = 'varBt_alt_bill_addr1'
            alt1_address_2nd_button_name = 'varBt_alt_bill_addr2'
            alt1_address_3rd_button_name = 'varBt_alt_bill_addr3'
            alt1_address_4th_button_name = 'varBt_alt_bill_addr4'
            alt1_address_5th_button_name = 'varBt_alt_bill_addr5'
            alt1_postcode_button_name = 'varBt_alt_bill_postcode'

            alt1_address1 = each_row['A1AAddress1']
            if alt1_address1 is None or alt1_address1.strip() == '':
                pass
            else:
                alt1_address_1st_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt1_address_1st_button_name))
                alt1_address_1st_button_element.send_keys(alt1_address1)

            alt1_address_2 = each_row['A1AAddress2']
            if alt1_address_2 is None or alt1_address_2.strip() == '':
                pass
            else:
                alt1_address_2nd_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt1_address_2nd_button_name))
                alt1_address_2nd_button_element.send_keys(alt1_address_2)

            alt1_address_3 = each_row['A1AAddress3']
            if alt1_address_3 is None or alt1_address_3.strip() == '':
                pass
            else:
                alt1_address_3rd_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt1_address_3rd_button_name))
                alt1_address_3rd_button_element.send_keys(alt1_address_3)

            alt1_address_4 = each_row['A1AAddress4']
            if alt1_address_4 is None or alt1_address_4.strip() == '':
                pass
            else:
                alt1_address_4th_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt1_address_4th_button_name))
                alt1_address_4th_button_element.send_keys(alt1_address_4)

            alt1_address_5 = each_row['A1AAddress5']
            if alt1_address_5 is None or alt1_address_5.strip() == '':
                pass
            else:
                alt1_address_5th_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt1_address_5th_button_name))
                alt1_address_5th_button_element.send_keys(alt1_address_5)

            if each_row['A1ACountry'] is None or each_row['A1ACountry'].strip() == '' \
                    or each_row['A1ACountry'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_alt_bill_cntry_cd'))
                select.select_by_value(each_row['A1ACountry'])

            if each_row['A1AState'] is None or each_row['A1AState'].strip() == '' \
                    or each_row['A1AState'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_alt_bill_state'))
                select.select_by_value(each_row['A1AState'])

            if each_row['A1ACity'] is None or each_row['A1ACity'].strip() == '' or \
                    each_row['A1AState'] != each_row['A1ACity'] or each_row['A1ACity'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_alt_bill_city'))
                select.select_by_value(each_row['A1ACity'])

            alt1_postcode = each_row['A1APostCode']
            if alt1_postcode is None or alt1_postcode.strip() == '':
                pass
            else:
                alt1_postcode_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt1_postcode_button_name))
                alt1_postcode_button_element.send_keys(alt1_postcode)
            # endregion

            # region Alternative-2 ADDRESS
            alt2_address_1st_button_name = 'varBt_alt2_bill_addr1'
            alt2_address_2nd_button_name = 'varBt_alt2_bill_addr2'
            alt2_address_3rd_button_name = 'varBt_alt2_bill_addr3'
            alt2_address_4th_button_name = 'varBt_alt2_bill_addr4'
            alt2_address_5th_button_name = 'varBt_alt2_bill_addr5'
            alt2_postcode_button_name = 'varBt_alt2_bill_postcode'

            alt2_address1 = each_row['A2AAddress1']
            if alt2_address1 is None or alt2_address1.strip() == '':
                pass
            else:
                alt2_address_1st_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt2_address_1st_button_name))
                alt2_address_1st_button_element.send_keys(alt2_address1)

            alt2_address2 = each_row['A2AAddress2']
            if alt2_address2 is None or alt2_address2.strip() == '':
                pass
            else:
                alt2_address_2nd_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt2_address_2nd_button_name))
                alt2_address_2nd_button_element.send_keys(alt2_address2)

            alt2_address3 = each_row['A2AAddress3']
            if alt2_address3 is None or alt2_address3.strip() == '':
                pass
            else:
                alt2_address_3rd_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt2_address_3rd_button_name))
                alt2_address_3rd_button_element.send_keys(alt2_address3)

            alt2_address4 = each_row['A2AAddress4']
            if alt2_address4 is None or alt2_address4.strip() == '':
                pass
            else:
                alt2_address_4th_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt2_address_4th_button_name))
                alt2_address_4th_button_element.send_keys(alt2_address4)

            alt2_address5 = each_row['A2AAddress5']
            if alt2_address5 is None or alt2_address5.strip() == '':
                pass
            else:
                alt2_address_5th_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt2_address_5th_button_name))
                alt2_address_5th_button_element.send_keys(alt2_address5)

            if each_row['A2ACountry'] is None or each_row['A2ACountry'].strip() == '' \
                    or each_row['A2ACountry'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_alt2_bill_cntry_cd'))
                select.select_by_value(each_row['A2ACountry'])

            if each_row['A2AState'] is None or each_row['A2AState'].strip() == '' \
                    or each_row['A2AState'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_alt2_bill_state'))
                select.select_by_value(each_row['A2AState'])

            if each_row['A2ACity'] is None or each_row['A2ACity'].strip() == '' or \
                    each_row['A2AState'] != each_row['A2ACity'] or each_row['A2ACity'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_alt2_bill_city'))
                select.select_by_value(each_row['A2ACity'])

            alt2_postcode = each_row['A2APostCode']
            if alt2_postcode is None or alt2_postcode.strip() == '':
                pass
            else:
                alt2_postcode_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(alt2_postcode_button_name))
                alt2_postcode_button_element.send_keys(alt2_postcode)
            # endregion

            # region COMPANY REFERENCES
            company_button_name = 'varBt_comp_name'
            cr_address1_button_name = 'varBt_comp_addr1'
            cr_address2_button_name = 'varBt_comp_addr2'
            cr_address3_button_name = 'varBt_comp_addr3'
            cr_address4_button_name = 'varBt_comp_addr4'
            cr_address5_button_name = 'varBt_comp_addr5'
            cr_telephone_button_name = 'varBt_comp_telno'
            cr_fax_button_name = 'varBt_comp_faxno'
            crdesignationbuttonname = 'varBt_comp_desgn'
            crservice_sincebuttonname = 'varBt_length_srv_year'
            basic_salary_anualbuttonname = 'varBt_mth_basic_salary'
            fix_allowancebuttonname = 'varBt_mth_fixed_allow'
            crpost_codebuttonname = 'varBt_comp_postcode'
            crother_incomebuttonname = 'varBt_other_income'

            company_name = each_row['CRCompanyName']
            if company_name is None or company_name.strip() == '':
                pass
            else:
                company_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(company_button_name))
                company_button_element.send_keys(company_name)

            cr_address1 = each_row['CRAddress1']
            if cr_address1 is None or cr_address1.strip() == '':
                pass
            else:
                cr_address1_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cr_address1_button_name))
                cr_address1_button_element.send_keys(cr_address1)

            cr_address2 = each_row['CRAddress2']
            if cr_address2 is None or cr_address2.strip() == '':
                pass
            else:
                cr_address2_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cr_address2_button_name))
                cr_address2_button_element.send_keys(cr_address2)

            cr_address3 = each_row['CRAddress3']
            if cr_address3 is None or cr_address3.strip() == '':
                pass
            else:
                cr_address3_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cr_address3_button_name))
                cr_address3_button_element.send_keys(cr_address3)

            cr_address4 = each_row['CRAddress4']
            if cr_address4 is None or cr_address4.strip() == '':
                pass
            else:
                cr_address4_button_element = WebDriverWait(driver, 1).until(
                    lambda dchromeriver: driver.find_element_by_name(cr_address4_button_name))
                cr_address4_button_element.send_keys(cr_address4)

            cr_address5 = each_row['CRAddress5']
            if cr_address5 is None or cr_address5.strip() == '':
                pass
            else:
                cr_address5_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cr_address5_button_name))
                cr_address5_button_element.send_keys(cr_address5)

            cr_telephone = each_row['CRTelephone']
            if cr_telephone is None or cr_telephone.strip() == '':
                pass
            else:
                cr_telephone_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cr_telephone_button_name))
                cr_telephone_button_element.send_keys(cr_telephone)

            cr_fax = each_row['CRFax']
            if cr_fax is None or cr_fax.strip() == '':
                pass
            else:
                cr_fax_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cr_fax_button_name))
                cr_fax_button_element.send_keys(cr_fax)

            cr_designation = each_row['CRDesignation']
            if cr_designation is None or cr_designation.strip() == '':
                pass
            else:
                cr_designation_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(crdesignationbuttonname))
                cr_designation_button_element.send_keys(cr_designation)

            if each_row['CRSelfEmployed'] is None or each_row['CRSelfEmployed'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_self_employ'))
                select.select_by_value(each_row['CRSelfEmployed'])

            cr_service_since = each_row['CRServiceSince']
            if cr_service_since is None or cr_service_since.strip() == '':
                pass
            else:
                cr_service_since = cr_service_since.split(' ', 1)
                cr_service_since_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(crservice_sincebuttonname))
                driver.find_element_by_name(crservice_sincebuttonname).clear()
                cr_service_since_button_element.send_keys(cr_service_since[0])

                cr_service_since_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name('varBt_length_srv_month'))
                driver.find_element_by_name('varBt_length_srv_month').clear()
                cr_service_since_button_element.send_keys(cr_service_since[1])

            if each_row['CRCountry'] is None or each_row['CRCountry'].strip() == '' \
                    or each_row['CRCountry'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_comp_cntry_cd'))
                select.select_by_value(each_row['CRCountry'])

            if each_row['CRStdIndustryCode'] is None or each_row['CRStdIndustryCode'].strip() == '' \
                    or each_row['CRStdIndustryCode'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_buss_sic'))
                select.select_by_value(each_row['CRStdIndustryCode'])

            if each_row['CRState'] is None or each_row['CRState'].strip() == '' \
                    or each_row['CRState'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_comp_state'))
                select.select_by_value(each_row['CRState'])

            basic_salary_anual = each_row['CRBasicSalaryAnnualIncome']
            if basic_salary_anual is None or basic_salary_anual.strip() == '':
                pass
            else:
                basic_salary_annual_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(basic_salary_anualbuttonname))
                basic_salary_annual_button_element.send_keys(basic_salary_anual)

            if each_row['CRCity'] is None or each_row['CRCity'].strip() == '' or \
                    each_row['CRState'] != each_row['CRCity'] or each_row['CRCity'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_comp_city'))
                select.select_by_value(each_row['CRCity'])

            fix_allowance = each_row['CRFixedAllowance']
            if fix_allowance is None or fix_allowance.strip() == '':
                pass
            else:
                fix_allowance_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(fix_allowancebuttonname))
                fix_allowance_button_element.send_keys(fix_allowance)

            crpost_code = each_row['CRPostCode']
            if crpost_code is None or crpost_code.strip() == '':
                pass
            else:
                crpost_code_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(crpost_codebuttonname))
                crpost_code_button_element.send_keys(crpost_code)

            crother_income = each_row['CROtherIncome']
            if crother_income is None or crother_income.strip() == '':
                pass
            else:
                crother_income_button_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(crother_incomebuttonname))
                crother_income_button_element.send_keys(crother_income)

            if each_row['CROrganisationType'] is None or each_row['CROrganisationType'].strip() == '' \
                    or each_row['CROrganisationType'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_comp_structure'))
                select.select_by_value(each_row['CROrganisationType'])

            if each_row['CRJobStabiliy'] is None or each_row['CRJobStabiliy'].strip() == '' \
                    or each_row['CRJobStabiliy'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_emp_stability'))
                select.select_by_value(each_row['CRJobStabiliy'])

            # endregion

            # region PREVIOUS COMPANY REFERENCES

            pre_employer_name = each_row['PCRPreviousEmployerName']
            pre_employname = 'varBt_ex_employ_name'
            if pre_employer_name is None or pre_employer_name.strip() == '':
                pass
            else:
                pre_employ_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(pre_employname))
                pre_employ_name_element.send_keys(pre_employer_name)

            pre_designation = each_row['PCRPreviousDesignation']
            pre_designationname = 'varBt_ex_employ_desgn'
            if pre_designation is None or pre_designation.strip() == '':
                pass
            else:
                pre_designation_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(pre_designationname))
                pre_designation_element.send_keys(pre_designation)

            pre_sl = each_row['PCRPreviousServiceLength']
            pre_slname = 'varBt_ex_employ_length_sev_year'
            pre_slmname = 'varBt_ex_employ_length_sev_month'

            if pre_sl.strip() == '' or pre_sl is None:
                pass
            else:
                pre_sl = pre_sl.split(' ', 1)
                pre_sl_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(pre_slname))
                driver.find_element_by_name(pre_slname).clear()
                pre_sl_name_element.send_keys(pre_sl[0])

                pre_slm_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(pre_slmname))
                driver.find_element_by_name(pre_slmname).clear()
                pre_slm_name_element.send_keys(pre_sl[1])

            # endregion

            # region BILLING AND CARD DELIVERY INFORMATION
            if each_row['PCRBillingCode'] is None or each_row['PCRBillingCode'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_bill_addr_cd'))
                select.select_by_value(each_row['PCRBillingCode'])

            if each_row['PCRCardCollectionMethod'] is None or each_row['PCRCardCollectionMethod'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_card_collection_cd'))
                select.select_by_value(each_row['PCRCardCollectionMethod'])

            if each_row['PCRLegalBillingAddressCode'] is None or \
                    each_row['PCRLegalBillingAddressCode'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_legal_addr_cd'))
                select.select_by_value(each_row['PCRLegalBillingAddressCode'])

            if each_row['PCRCardDeliveryAddressCode'] is None or \
                    each_row['PCRCardDeliveryAddressCode'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_card_delivery_addr_cd'))
                select.select_by_value(each_row['PCRCardDeliveryAddressCode'])

            if each_row['PCRSeparatedStatement'] is None or each_row['PCRSeparatedStatement'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_separate_stmt_ind'))
                select.select_by_value(each_row['PCRSeparatedStatement'])
            # endregion

            # region OTHER LOANS
            loan_bank_1 = each_row['OLLoanBank1']
            loan_bank_1name = 'varBt_oth_loan_bank_1'
            if loan_bank_1 is None or loan_bank_1.strip() == '':
                pass
            else:
                loan_bank1_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_1name))
                loan_bank1_name_element.send_keys(loan_bank_1)

            loan_bank_type1 = each_row['OLType1']
            loan_bank_type1name = 'varBt_oth_loan_type_1'
            if loan_bank_type1 is None or loan_bank_type1.strip() == '':
                pass
            else:
                loan_bank_type1_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_type1name))
                loan_bank_type1_element.send_keys(loan_bank_type1)

            loan_bank_amount1 = each_row['OLAmount1']
            loan_bank_amount1name = 'varBt_oth_loan_amt_1'
            if loan_bank_amount1 is None or loan_bank_amount1.strip() == '':
                pass
            else:
                loan_bank_amount1_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_amount1name))
                loan_bank_amount1_element.send_keys(loan_bank_amount1)

            loan_bank_2 = each_row['OLLoanBank2']
            loan_bank_2name = 'varBt_oth_loan_bank_2'
            if loan_bank_2 is None or loan_bank_2.strip() == '':
                pass
            else:
                loan_bank2_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_2name))
                loan_bank2_name_element.send_keys(loan_bank_2)

            loan_bank_type2 = each_row['OLType2']
            loan_bank_type2name = 'varBt_oth_loan_type_2'
            if loan_bank_type2 is None or loan_bank_type2.strip() == '':
                pass
            else:
                loan_bank_type2_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_type2name))
                loan_bank_type2_element.send_keys(loan_bank_type2)

            loan_bank_amount2 = each_row['OLAmount2']
            loan_bank_amount2name = 'varBt_oth_loan_amt_2'
            if loan_bank_amount2 is None or loan_bank_amount2.strip() == '':
                pass
            else:
                loan_bank_amount2_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_amount2name))
                loan_bank_amount2_element.send_keys(loan_bank_amount2)

            loan_bank_3 = each_row['OLLoanBank3']
            loan_bank_3name = 'varBt_oth_loan_bank_3'
            if loan_bank_3 is None or loan_bank_3.strip() == '':
                pass
            else:
                loan_bank3_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_3name))
                loan_bank3_name_element.send_keys(loan_bank_3)

            loan_bank_type3 = each_row['OLType3']
            loan_bank_type3name = 'varBt_oth_loan_type_3'
            if loan_bank_type3 is None or loan_bank_type3.strip() == '':
                pass
            else:
                loan_bank_type3_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_type3name))
                loan_bank_type3_element.send_keys(loan_bank_type3)

            loan_bank_amount3 = each_row['OLAmount3']
            loan_bank_amount3name = 'varBt_oth_loan_amt_3'
            if loan_bank_amount3 is None or loan_bank_amount3.strip() == '':
                pass
            else:
                loan_bank_amount2_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(loan_bank_amount3name))
                loan_bank_amount2_element.send_keys(loan_bank_amount3)

            card_no1 = each_row['OLCardNo1']
            card_no1name = 'varBt_cr_card_no_11'
            if card_no1 is None or card_no1.strip() == '':
                pass

            else:
                card_no1_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(card_no1name))
                card_no1_name_element.send_keys(card_no1)

            card_limit1 = each_row['OLCardLimit1']
            card_limit1name = 'varBt_cr_card_limit_1'
            if card_limit1 is None or card_limit1.strip() == '':
                pass
            else:
                card_limit1_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(card_limit1name))
                card_limit1_name_element.send_keys(card_limit1)

            if each_row['OLBalanceTransfer1'] is None or each_row['OLBalanceTransfer1'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_bal_transfer_ind_1'))
                select.select_by_value(each_row['OLBalanceTransfer1'])

            card_no2 = each_row['OLCardNo2']
            card_no2name = 'varBt_cr_card_no_21'
            if card_no2 is None or card_no2.strip() == '':
                pass
            else:
                card_no2_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(card_no2name))
                card_no2_name_element.send_keys(card_no2)

            card_limit2 = each_row['OLCardLimit2']
            card_limit2name = 'varBt_cr_card_limit_2'
            if card_limit2 is None or card_limit2.strip() == '':
                pass
            else:
                card_limit2_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(card_limit2name))
                card_limit2_name_element.send_keys(card_limit2)

            if (each_row['OLBalanceTransfer2']) is None or each_row['OLBalanceTransfer2'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_bal_transfer_ind_2'))
                select.select_by_value(each_row['OLBalanceTransfer2'])

            card_no3 = each_row['OLCardNo3']
            card_no3name = 'varBt_cr_card_no_31'
            if card_no3 is None or card_no3.strip() == '':
                pass
            else:
                card_no3_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(card_no3name))
                card_no3_name_element.send_keys(card_no3)

            card_limit3 = each_row['OLCardLimit3']
            card_limit3name = 'varBt_cr_card_limit_3'
            if card_limit3 is None or card_limit3.strip() == '':
                pass
            else:
                card_limit3_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(card_limit3name))
                card_limit3_name_element.send_keys(card_limit3)

            if each_row['OLBalanceTransfer3'] is None or each_row['OLBalanceTransfer3'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_bal_transfer_ind_3'))
                select.select_by_value(each_row['OLBalanceTransfer3'])

            if each_row['OLCreditStatus'] is None or each_row['OLCreditStatus'].strip() == '' \
                    or each_row['OLCreditStatus'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_oth_bank_cr_status'))
                select.select_by_value(each_row['OLCreditStatus'])
            # endregion

            # region SPOUSE REFERENCES
            spousename = each_row['SRSpouseName']
            spousebuttonname = 'varBt_spouse_name'
            if spousename is None or spousename.strip() == '':
                pass
            else:
                spouse_button_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(spousebuttonname))
                spouse_button_name_element.send_keys(spousename)

            spouse_nricno = each_row['SRSpouseNRICNumber']
            spouse_nricnoname = 'varBt_spouse_nric'
            if spouse_nricno is None or spouse_nricno.strip() == '':
                pass
            else:
                spouse_nricno_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(spouse_nricnoname))
                spouse_nricno_name_element.send_keys(spouse_nricno)

            semployer_name = each_row['SREmployerName']
            semployerbuttonname = 'varBt_spouse_employ_name'
            if semployer_name is None or semployer_name.strip() == '':
                pass
            else:
                s_employer_button_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(semployerbuttonname))
                s_employer_button_name_element.send_keys(semployer_name)
            # endregion

            # region Contact Person Detail
            contact_name = each_row['CPDContactName']
            c_name = 'varBt_rel_name'
            if contact_name is None or contact_name.strip() == '':
                pass
            else:
                c_name_element = WebDriverWait(driver, 1).\
                    until(lambda chromedriver: driver.find_element_by_name(c_name))
                c_name_element.send_keys(contact_name)

            cp_nricno = each_row['CPDNRICNumber']
            cp_nricnoname = 'varBt_rel_nric'
            if cp_nricno is None or cp_nricno.strip() == '':
                pass
            else:
                cp_nricno_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cp_nricnoname))
                cp_nricno_name_element.send_keys(cp_nricno)

            cpd_home_address = each_row['CPDHomeAddress']
            cpd_home_addressname = 'varBt_rel_addr1'
            if cpd_home_address is None or cpd_home_address.strip() == '':
                pass
            else:
                cpd_home_address_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cpd_home_addressname))
                cpd_home_address_element.send_keys(cpd_home_address)

            if each_row['CPDGender'] is None or each_row['CPDGender'].strip() == '' \
                    or each_row['CPDGender'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_rel_sex'))
                select.select_by_visible_text(each_row['CPDGender'])

            cp_postcode = each_row['CPDPostCode']
            cp_postcodename = 'varBt_rel_addr2'
            if cp_postcode is None or cp_postcode.strip() == '':
                pass
            else:
                cp_postcode_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cp_postcodename))
                cp_postcode_name_element.send_keys(cp_postcode)

            cpd_dob = each_row['CPDDateOfBirth']
            cpd_dobname = 'varBt_rel_dob'
            if cpd_dob is None or cpd_dob.strip() == '':
                pass
            else:
                cpd_dobname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cpd_dobname))
                cpd_dobname_element.send_keys(cpd_dob)

            cpd_mob = each_row['CPDMobile']
            cpd_mobname = 'varBt_rel_addr3'
            if cpd_mob is None or cpd_mob.strip() == '':
                pass
            else:
                cpd_mob_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cpd_mobname))
                cpd_mob_element.send_keys(cpd_mob)

            cpd_tel = each_row['CPDTelephone']
            cpd_telname = 'varBt_rel_addr3'
            if cpd_tel is None or cpd_tel.strip() == '':
                pass
            else:
                cpd_telname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cpd_telname))
                cpd_telname_element.send_keys(cpd_tel)

            cpd_comp = each_row['CPDCompanyName']
            cpd_compname = 'varBt_rel_addr4'
            if cpd_comp is None or cpd_comp.strip() == '':
                pass
            else:
                cpd_compname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cpd_compname))
                cpd_compname_element.send_keys(cpd_comp)

            if each_row['CPDRelationship'] is None or each_row['CPDRelationship'].strip() == '' \
                    or each_row['CPDRelationship'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_rel_relation_to_ch'))
                select.select_by_value(each_row['CPDRelationship'])

            cpd_desig = each_row['CPDDesignation']
            cpd_designame = 'varBt_rel_desgn'
            if cpd_desig is None or cpd_desig.strip() == '':
                pass
            else:
                cpd_designame_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(cpd_designame))
                cpd_designame_element.send_keys(cpd_desig)

            if each_row['CPDState'] is None or each_row['CPDState'].strip() == '' \
                    or each_row['CPDState'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_rel_state'))
                select.select_by_value(each_row['CPDState'])

            if each_row['CPDCountry'] is None or each_row['CPDCountry'].strip() == '' \
                    or each_row['CPDCountry'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_rel_cntry_cd'))
                select.select_by_value(each_row['CPDCountry'])

            if each_row['CPDCity'] is None or each_row['CPDCity'].strip() == '' or \
                    each_row['CPDState'] != each_row['CPDCity'] or each_row['CPDCity'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_rel_city'))
                select.select_by_value(each_row['CPDCity'])
            # endregion

            # region OTHERS
            orecom_name = each_row['ORecommenderName']
            orecommender_name = 'varBt_recom_card_name'
            if orecom_name is None or orecom_name.strip() == '':
                pass
            else:
                orecommender_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(orecommender_name))
                orecommender_name_element.send_keys(orecom_name)

            orecom_pay_aac = each_row['ORecommenderPaymentAccount']
            orecom_pay_aacno = 'varBt_rwds_accno'
            if orecom_pay_aac is None or orecom_pay_aac.strip() == '':
                pass
            else:
                orecom_pay_aacno_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(orecom_pay_aacno))
                orecom_pay_aacno_element.send_keys(orecom_pay_aac)

            orecom_number = each_row['ORecommenderNumber']
            orecommender_number = 'varBt_recom_card_no'
            if orecom_number is None or orecom_number.strip() == '':
                pass
            else:
                orecommender_number_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(orecommender_number))
                orecommender_number_element.send_keys(orecom_number)

            if each_row['OCurrentBankCustomer'] is None or each_row['OCurrentBankCustomer'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_curr_bank_rel'))
                select.select_by_value(each_row['OCurrentBankCustomer'])

            oreq_climit = each_row['ORequestedCreditLimit']
            oreq_climit_name = 'varBt_request_credit_limit1'
            if oreq_climit is None or oreq_climit.strip() == '':
                pass
            else:
                oreq_climit_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(oreq_climit_name))
                oreq_climit_name_element.send_keys(oreq_climit)

            if each_row['OCustomerClass'] is None or each_row['OCustomerClass'].strip() == '' \
                    or each_row['OCustomerClass'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_cust_class'))
                select.select_by_value(each_row['OCustomerClass'])

            if each_row['OInterestGroup'] is None or each_row['OInterestGroup'].strip() == '' \
                    or each_row['OInterestGroup'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_interest_grp_code'))
                select.select_by_value(each_row['OInterestGroup'])

            if each_row['OCreditGuarantee'] is None or each_row['OCreditGuarantee'].strip() == '' \
                    or each_row['OCreditGuarantee'] == 'NS':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_credit_guarantee'))
                select.select_by_value(each_row['OCreditGuarantee'])

            if each_row['OPhotoIndicator'] is None or each_row['OPhotoIndicator'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_photo_ind'))
                select.select_by_value(each_row['OPhotoIndicator'])

            if each_row['ODocumentComplete'] is None or each_row['ODocumentComplete'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_doc_complete'))
                select.select_by_value(each_row['ODocumentComplete'])

            if each_row['OLanguageIndicator'] is None or each_row['OLanguageIndicator'].strip() == '':
                pass
            else:
                select = Select(driver.find_element_by_name('varBt_language_ind'))
                select.select_by_value(each_row['OLanguageIndicator'])

            select = Select(driver.find_element_by_name('varBt_pin_gen_ind'))
            select.select_by_value(each_row['OATMIndicator'])

            # endregion

            # region autopay indicator
            select = Select(driver.find_element_by_name('varBt_autopay_ind'))
            if each_row['OAutopayIndicator'].strip() != '' or each_row['OAutopayIndicator'] is not None:

                select.select_by_value(each_row['OAutopayIndicator'])
                try:
                    if each_row['OAutopayIndicator'].strip() == '3' or each_row['OAutopayIndicator'].strip() == '4' or \
                            each_row['OAutopayIndicator'].strip() == '5':
                        if each_row['LADLocalAutopayBranch'] == 'NS':
                            pass
                        else:
                            select = Select(driver.find_element_by_name('varBt_local_autopay_bank_branch'))
                            select.select_by_value(each_row['LADLocalAutopayBranch'])

                        local_autopay_accnt = each_row['LADLocalAutopayAccount']
                        local_autopay_accntbuttonname = 'varBt_local_autopay_bank_accno'
                        local_autopay_accntbutton_element = WebDriverWait(driver, 1).until(
                            lambda chromedriver: driver.find_element_by_name(local_autopay_accntbuttonname))
                        local_autopay_accntbutton_element.send_keys(local_autopay_accnt)

                        local_autopay_accnt_name = each_row['LADLocalAutopayAccountName']
                        local_autopay_accnt_namebuttonname = 'varBt_local_autopay_accno_name'
                        local_autopay_accnt_namebutton_element = WebDriverWait(driver, 1).until(
                            lambda chromedriver: driver.find_element_by_name(local_autopay_accnt_namebuttonname))
                        local_autopay_accnt_namebutton_element.send_keys(local_autopay_accnt_name)

                        if each_row['OAutopayIndicator'].strip() == '5':
                            local_autopay_rate = each_row['LADLocalAutopayRate']
                            local_autopay_ratebuttonname = 'varBt_autopay_rate'
                            local_autopay_ratebutton_element = WebDriverWait(driver, 1).until(
                                lambda chromedriver: driver.find_element_by_name(local_autopay_ratebuttonname))
                            local_autopay_ratebutton_element.send_keys(local_autopay_rate)
                except:
                    if each_row['OAutopayIndicator'].strip() == '3' or each_row['OAutopayIndicator'].strip() == '4' or \
                            each_row['OAutopayIndicator'].strip() == '5':
                        if each_row['FADForeignAutopayBranch'] == 'NS':
                            pass
                        else:
                            select = Select(driver.find_element_by_name('varBt_frgn_autopay_bank_branch'))
                            select.select_by_value(each_row['FADForeignAutopayBranch'])

                        local_autopay_accnt = each_row['FADForeignAutopayAccount']
                        local_autopay_accntbuttonname = 'varBt_frgn_autopay_bank_accno'
                        local_autopay_accntbutton_element = WebDriverWait(driver, 1).until(
                            lambda chromedriver: driver.find_element_by_name(local_autopay_accntbuttonname))
                        local_autopay_accntbutton_element.send_keys(local_autopay_accnt)

                        local_autopay_accnt_name = each_row['FADForeignAutopayAccountName']
                        local_autopay_accnt_namebuttonname = 'varBt_frgn_autopay_accno_name'
                        local_autopay_accnt_namebutton_element = WebDriverWait(driver, 1).until(
                            lambda chromedriver: driver.find_element_by_name(local_autopay_accnt_namebuttonname))
                        local_autopay_accnt_namebutton_element.send_keys(local_autopay_accnt_name)

                        if each_row['OAutopayIndicator'].strip() == '5':
                            local_autopay_rate = each_row['FADForeignAutopayRate']
                            local_autopay_ratebuttonname = 'varBt_frgn_autopay_rate'
                            local_autopay_ratebutton_element = WebDriverWait(driver, 1).until(
                                lambda chromedriver: driver.find_element_by_name(local_autopay_ratebuttonname))
                            local_autopay_ratebutton_element.send_keys(local_autopay_rate)

            # endregion

            # region ADDITIONAL
            knext_rev = each_row['AIDKycNextReviewDate']
            knext_rev_name = 'varCb_adl_field_date10'
            if knext_rev is None or knext_rev.strip() == '':
                pass
            else:
                knext_rev_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(knext_rev_name))
                knext_rev_name_element.send_keys(knext_rev)

            krisk = each_row['AINKycRiskInd']
            krisk_name = 'varCb_adl_field_num01'
            if krisk is None or krisk.strip() == '':
                pass
            else:
                krisk_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(krisk_name))
                krisk_name_element.send_keys(krisk)

            # charcter
            fathername = each_row['father']
            father_name = 'varCb_adl_field_name01'
            if fathername is None or fathername.strip() == '':
                pass
            else:
                father_name_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(father_name))
                father_name_element.send_keys(fathername)

            nid = each_row['NID']
            nid_number = 'varCb_adl_field_name02'
            if nid is None or nid.strip() == '':
                pass
            else:
                nid_number_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(nid_number))
                nid_number_element.send_keys(nid)

            app_serial = each_row['AICApplicationSerialNo']
            app_serial_no = 'varCb_adl_field_name03'
            if app_serial is None or app_serial.strip() == '':
                pass
            else:
                app_serial_no_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(app_serial_no))
                app_serial_no_element.send_keys(app_serial)

            sector_code = each_row['AICSectorCode']
            sector_codename = 'varCb_adl_field_name04'
            if sector_code is None or sector_code.strip() == '':
                pass
            else:
                sector_codename_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(sector_codename))
                sector_codename_element.send_keys(sector_code)

            tin_no = each_row['TinNo']
            tin_noname = 'varCb_adl_field_name05'
            if tin_no is None or tin_no.strip() == '':
                pass
            else:
                tin_noname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(tin_noname))
                tin_noname_element.send_keys(tin_no)

            secured_card = each_row['AICLienAmtAgainstSecuredCard']
            secured_cardname = 'varCb_adl_field_name06'
            if secured_card is None or secured_card.strip() == '':
                pass
            else:
                secured_cardname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(secured_cardname))
                secured_cardname_element.send_keys(secured_card)

            priority_pass = each_row['AICPriorityPass']
            priority_passname = 'varCb_adl_field_name07'
            if priority_pass is None or priority_pass.strip() == '':
                pass
            else:
                priority_passname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(priority_passname))
                priority_passname_element.send_keys(priority_pass)

            passport_detail = each_row['AICPassportDetail']
            passport_detailname = 'varCb_adl_field_name08'
            if passport_detail is None or passport_detail.strip() == '':
                pass
            else:
                passport_detailname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(passport_detailname))
                passport_detailname_element.send_keys(passport_detail)

            passport_issue = each_row['AICPassportIssued']
            passport_issuename = 'varCb_adl_field_name09'
            if passport_issue is None or passport_issue.strip() == '':
                pass
            else:
                passport_issuename_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(passport_issuename))
                passport_issuename_element.send_keys(passport_issue)

            passport_expiry = each_row['AICPassportExpiry']
            passport_expiryname = 'varCb_adl_field_name10'
            if passport_expiry is None or passport_expiry.strip() == '':
                pass
            else:
                passport_expiryname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(passport_expiryname))
                passport_expiryname_element.send_keys(passport_expiry)

            passport_issue_place = each_row['AICPassportIssuePlace']
            passport_issue_placename = 'varCb_adl_field_name11'
            if passport_issue_place is None or passport_issue_place.strip() == '':
                pass
            else:
                passport_issue_placename_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(passport_issue_placename))
                passport_issue_placename_element.send_keys(passport_issue_place)

            srecom_number = each_row['AICSuppleRecommenderNumber']
            srecom_numbername = 'varCb_adl_field_name13'
            if srecom_number is None or srecom_number.strip() == '':
                pass
            else:
                srecom_numbername_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(srecom_numbername))
                srecom_numbername_element.send_keys(srecom_number)

            com_type = each_row['AICCompanyType']
            com_typename = 'varCb_adl_field_name14'
            if com_type is None or com_type.strip() == '':
                pass
            else:
                com_typename_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(com_typename))
                com_typename_element.send_keys(com_type)

            app_level = each_row['AICApprovalLevel']
            app_levelname = 'varCb_adl_field_name17'
            if app_level is None or app_level.strip() == '':
                pass
            else:
                app_levelname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(app_levelname))
                app_levelname_element.send_keys(app_level)

            app_pin = each_row['AICApproverPIN']
            app_pinname = 'varCb_adl_field_name18'
            if app_pin is None or app_pin.strip() == '':
                pass
            else:
                app_pinname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(app_pinname))
                app_pinname_element.send_keys(app_pin)

            analyst_pin = each_row['AICAnalystPIN']
            analyst_pinname = 'varCb_adl_field_name19'
            if analyst_pin is None or analyst_pin.strip() == '':
                pass
            else:
                analyst_pinname_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(analyst_pinname))
                analyst_pinname_element.send_keys(analyst_pin)

            ss_code = each_row['AICSubSectorCodeSSS']
            ss_codename = 'varCb_adl_field_name20'
            if ss_code is None or ss_code.strip() == '':
                pass
            else:
                ss_codename_element = WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_name(ss_codename))
                ss_codename_element.send_keys(ss_code)
            # endregion

            # clicking add button of form
            add2_button_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[2]/tbody/tr/td/input[1]'
            WebDriverWait(driver, 1).\
                until(lambda chromedriver: driver.find_element_by_xpath(add2_button_xpath)).click()
            driver.switch_to.alert.accept()

            # region NEXT PAGE

            select = Select(driver.find_element_by_name('varBt_decision_code'))
            select.select_by_visible_text('1-Approved')

            try:
                l_credit = 'varBt_local_credit_limit'
                f_credit = 'varBt_foreign_credit_limit'

                try:
                    lc_visible = driver.find_element_by_name(l_credit).is_displayed()
                except:
                    lc_visible = False

                try:
                    fc_visible = driver.find_element_by_name(f_credit).is_displayed()
                except:
                    fc_visible = False

                credit_limit = each_row['PRDCreditLimit']
                loan_id = each_row['loanID']
                # BasicLoanID = loanId

                if lc_visible is True and fc_visible is True:
                    creditlimit1_element = WebDriverWait(driver, 1).until(
                        lambda chromedriver: driver.find_element_by_name(l_credit))
                    creditlimit1_element.send_keys(credit_limit)

                else:
                    if lc_visible is True:
                        creditlimit1_element = WebDriverWait(driver, 1).until(
                            lambda chromedriver: driver.find_element_by_name(l_credit))
                        creditlimit1_element.send_keys(credit_limit)

                    else:
                        if fc_visible is True:
                            creditlimit2_element = WebDriverWait(driver, 1).until(
                                lambda chromedriver: driver.find_element_by_name(f_credit))
                            creditlimit2_element.send_keys(credit_limit)

                # TODO:Reactive Date

                if each_row['NPBillingCycle'] is None or each_row['NPBillingCycle'].strip() == '':
                    pass
                else:
                    select = Select(driver.find_element_by_name('varBt_bill_cycle'))
                    select.select_by_value(each_row['NPBillingCycle'])

                if each_row['NPMonitorCode'] is None or each_row['NPMonitorCode'].strip() == '':
                    pass
                else:
                    select = Select(driver.find_element_by_name('varBt_monitor_code'))
                    select.select_by_value(each_row['NPMonitorCode'])
                # TODO stopdunning code

                select = Select(driver.find_element_by_name('varBt_stop_dunn_code'))
                select.select_by_visible_text('N-No')

                if each_row['NPDirectMailIndicator'] is None or each_row['NPDirectMailIndicator'].strip() == '':
                    pass
                else:
                    select = Select(driver.find_element_by_name('varBt_direct_mail_ind'))
                    select.select_by_value(each_row['NPDirectMailIndicator'])

                application_no = driver.find_element_by_xpath(
                    "/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[2]/td[2]").text
                print(application_no)

                add3_button_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[2]/tbody/tr/td/input[1]'
                WebDriverWait(driver, 1).until(
                    lambda chromedriver: driver.find_element_by_xpath(add3_button_xpath)).click()
                # time.sleep(10) # for testing purpose
                driver.switch_to.alert.accept()
                # time.sleep(1)
                driver.switch_to.alert.accept()
                # time.sleep(1)

                is_true = False
                while is_true is False:
                    is_true = basic_limit_visible(driver, logging)

                    if is_true is True:
                        raise Exception
                    else:
                        break

            except Exception as e:
                logging.warning('_From-> BasicCard->_ :' + str(e) + "\n" + each_row['loanID'])
                loan_id = each_row['loanID']
                procedure_name = "EXEC RDMS_UpdateCardCreditError_Card_Rpa @loanId=?"
                object_connection = databaseConnection()
                object_connection.callProcedureError(procedure_name, connection_string, loan_id)
                limit_raise = True
                raise Exception

            # endregion

            # skip button click

            skip_button_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[4]/tbody/tr/td/input[2]'
            WebDriverWait(driver, 1).until(lambda chromedriver: driver.find_element_by_xpath(skip_button_xpath)).click()

            procedure_name = "EXEC RDMS_UpdateApplicationNo_Card_Rpa @loanId=?, @ApplicationNo=?"
            object_connection = databaseConnection()
            object_connection.callUpdateProcedure(procedure_name, connection_string, loan_id, application_no)

        except Exception as e:
            logging.warning('_From-> BasicCard->_ :' + str(e) + "\n" + each_row['loanID'])
            if limit_raise:
                limit_raise = False
                raise Exception

            else:
                loan_id = each_row['loanID']
                procedure_name = "EXEC RDMS_UpdateCardNoError_Card_Rpa @loanId=?"
                object_connection = databaseConnection()
                object_connection.callProcedureError(procedure_name, connection_string, loan_id)
                raise Exception
            # return Exception

        # supply connect

        results = []
        procedure_name = "EXEC RDMS_GetCreditCardInfoBySerialNo_SupplementaryCard_Rpa_py @BasicLoanId=?"
        object_connection = databaseConnection()
        results = object_connection.callProcedureSupply(procedure_name, connection_string, loan_id)

        # driver.switch_to.alert.accept()
        # driver.switch_to.alert.accept()

        if results.__len__() == 0:
            driver.switch_to.alert.dismiss()
            # click return button
            return_button_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[3]/tbody/tr/td/input[2]'
            WebDriverWait(driver, 1).\
                until(lambda chromedriver: driver.find_element_by_xpath(return_button_xpath)).click()

        else:
            while results.__len__() == 1:
                try:
                    object_supply = Supply()
                    results = object_supply.SupplyFields(results, driver, loan_id, connection_string, logging)
                    results = []
                    procedure_name = "EXEC RDMS_GetCreditCardInfoBySerialNo_SupplementaryCard_Rpa_py @BasicLoanId=?"
                    object_connection = databaseConnection()
                    results = object_connection.callProcedureSupply(procedure_name, connection_string, loan_id)

                except Exception as e:
                    logging.warning('_From-> BasicCard->_ :' + str(e) + '\n supplementary search error')
                    results = []
                    procedure_name = "EXEC RDMS_GetCreditCardInfoBySerialNo_SupplementaryCard_Rpa_py @BasicLoanId=?"
                    object_connection = databaseConnection()
                    results = object_connection.callProcedureSupply(procedure_name, connection_string, loan_id)

            driver.switch_to.alert.dismiss()
            # click return button
            return_button_xpath = '/html/body/table[3]/tbody/tr/td[3]/form/table[3]/tbody/tr/td/input[2]'
            WebDriverWait(driver, 1).\
                until(lambda chromedriver: driver.find_element_by_xpath(return_button_xpath)).click()
