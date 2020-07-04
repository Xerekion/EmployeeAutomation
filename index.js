const {Builder, By, Key, until} = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome')
const options = new chrome.Options();
//const chromeCapabilities = new chrome.chromeCapabilities();
options.setUserPreferences( { 'download.default_directory': 'F:\Selenium\seleniumExample\download' })
options.setUserPreferences({"download.prompt_for_download": false});
options.setUserPreferences({"download.directory_upgrade": true});
options.setUserPreferences({"safebrowsing.enabled": true});
options.addArguments('--test-type', '--start-maximized');

var xlsx = require('xlsx');
var wb = xlsx.readFile("EmployeeData.xls",{type:'binary', cellDates:true, cellNF: false, cellText:false});

var ws = wb.Sheets["Sheet1"]

var data = xlsx.utils.sheet_to_json(ws, {defval: " "});



function GenerateData(){
    for(u = 0; u < data.length; u++){
        var randomString = "";
        for(i = 0; i < 9; i++){
            var randomNum = Math.floor(Math.random() * 127) + 33;
            var randomChar = String.fromCharCode(randomNum);
            randomString = randomString + randomChar.toString();
        }
        data[u].Username = data[u].First_Name + randomString;
    }
}
GenerateData();

console.log(data)
//var newWB = xlsx.utils.book_new();

wb.Sheets["Sheet1"] = xlsx.utils.json_to_sheet(data);
//xlsx.utils.book_append_sheet(wb,newWS);
//console.log(wb.Sheets["Sheet1"]);
xlsx.writeFile(wb,"EmployeeData.xls",{cellDates:true});

//const driver = new Builder().forBrowser('chrome').setChromeOptions(options).build();
//async function login(){
//await driver.get('https://opensource-demo.orangehrmlive.com');
//await driver.findElement(By.name("txtUsername")).sendKeys("Admin");
//await driver.findElement(By.name("txtPassword")).sendKeys("admin123", Key.RETURN);
//await driver.get('https://opensource-demo.orangehrmlive.com/index.php/pim/addEmployee');
//await driver.findElement(By.name("chkLogin")).click();
//}
//login();
