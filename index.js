const {Builder, By, Key, until} = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome')
const options = new chrome.Options();
//const chromeCapabilities = new chrome.chromeCapabilities();
options.setUserPreferences( { 'download.default_directory': 'F:\Selenium\seleniumExample\download' })
options.setUserPreferences({"download.prompt_for_download": false});
options.setUserPreferences({"download.directory_upgrade": true});
options.setUserPreferences({"safebrowsing.enabled": true});
options.addArguments('--test-type', '--start-maximized');

const driver = new Builder().forBrowser('chrome').setChromeOptions(options).build();
async function login(){
await driver.get('https://opensource-demo.orangehrmlive.com');
await driver.findElement(By.name("txtUsername")).sendKeys("Admin");
await driver.findElement(By.name("txtPassword")).sendKeys("admin123", Key.RETURN);
await driver.get('https://opensource-demo.orangehrmlive.com/index.php/pim/addEmployee');
await driver.findElement(By.name("chkLogin")).click();
}
login();
