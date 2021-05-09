from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec


def supply_limit_visible(driver, logging):
    """

    :param logging:
    :param driver:
    :return:
    """
    try:
        WebDriverWait(driver, 7).until(
            ec.presence_of_element_located(
                (By.XPATH, "/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[2]/td[2]"))
        )
        return True

    except Exception as e:
        logging.warning('_From-> supply_limit_visible()_ :' + str(e))
        return False


def basic_limit_visible(driver, logging):
    try:
        WebDriverWait(driver, 5).until(
            ec.presence_of_element_located(
                (By.XPATH, "/html/body/table[3]/tbody/tr/td[3]/form/table[1]/tbody/tr[2]/td[2]"))
        )
        return True

    except Exception as e:
        logging.warning('_From-> basic_limit_visible()_ :' + str(e))
        return False
