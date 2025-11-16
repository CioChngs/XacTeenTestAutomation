import pytest
from appium import webdriver
from appium.options.android import UiAutomator2Options

@pytest.fixture
def driver():
    options = UiAutomator2Options()
    options.set_capability("platformName", "Android")
    options.set_capability("automationName", "UiAutomator2")
    options.set_capability("appPackage", "mn.xacbank.teen")
    options.set_capability("appActivity", "mn.xacbank.teen.MainActivity")
    options.set_capability("noReset", True)
    options.set_capability("uiautomator2ServerLaunchTimeout", 60000)
    driver = webdriver.Remote("http://127.0.0.1:4723", options=options)
    yield driver
    driver.quit()
