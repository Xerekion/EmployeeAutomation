const {Builder, By, Key, until, WebElement} = require('selenium-webdriver');
const firefox = require('selenium-webdriver/firefox')
const options = new firefox.Options();
//const chromeCapabilities = new chrome.chromeCapabilities();
//options.setUserPreferences( { 'download.default_directory': 'F:\Selenium\seleniumExample\download' })
//options.setUserPreferences({"download.prompt_for_download": false});
//options.setUserPreferences({"download.directory_upgrade": true});
//options.setUserPreferences({"safebrowsing.enabled": true});
//options.addArguments('--test-type', '--start-maximized');
//options.setPreference()

var xlsx = require('xlsx');
var wb = xlsx.readFile("EmployeeData.xls",{type:'binary', cellDates:true, cellNF: false, cellText:false});

var ws = wb.Sheets["Sheet1"]

var data = xlsx.utils.sheet_to_json(ws, {defval: " ",raw:false,dateNF:'yyyy-mm-dd'});




function GenerateLoginData(){
    for(u = 0; u < data.length; u++){
        if(data[u].Is_Created == "Yes"){
            continue;
        }
        else{
            var randomString = "";
            randomString = data[u].First_Name;
            if(randomString.charAt(0) != randomString.charAt(0).toUpperCase()){
                randomString = randomString.replace(randomString.charAt(0) ,randomString.charAt(0).toUpperCase());
                data[u].First_Name = randomString;
            }
            randomString = data[u].Last_Name;
            if(randomString.charAt(0) != randomString.charAt(0).toUpperCase()){
                randomString = randomString.replace(randomString.charAt(0) ,randomString.charAt(0).toUpperCase());
                data[u].Last_Name = randomString;
            }
            if(data[u].Username == " " || undefined){
                for(i = 0; i < 5; i++){
                    var randomNum = Math.floor(Math.random() * 26) + 97;
                    var randomChar = String.fromCharCode(randomNum);
                    randomString = randomString + randomChar.toString();
                }
                data[u].Username = data[u].First_Name + randomString;
                randomString = "";
            }
            if(data[u].Password == " " || undefined){
                for(i = 0; i < 5; i++){
                    var randomNum = Math.floor(Math.random() * 10);
                    //var randomChar = String.fromCharCode(randomNum);
                    randomString = randomString + randomNum.toString();
                }
                for(i = 0; i < 3; i++){
                    var randomNum = Math.floor(Math.random() * 15) + 33;
                    var randomChar = String.fromCharCode(randomNum);
                    randomString = randomString + randomChar.toString();
                }
                data[u].Password = data[u].Last_Name + randomString;
            }
            if(data[u].Employee_id == " " || undefined){
                data[u].Employee_id = "testid:" + u.toString();
            }
        }
    }
}
GenerateLoginData();

//console.log(today);


wb.Sheets["Sheet1"] = xlsx.utils.json_to_sheet(data);

xlsx.writeFile(wb,"EmployeeData.xls",{cellDates:true});

//const driver = new Builder().forBrowser('chrome').setChromeOptions(options).build();
const driver = new Builder().forBrowser('firefox').setFirefoxOptions(options).build();
driver.manage().window().maximize();
async function login(){
    await driver.get('https://opensource-demo.orangehrmlive.com');
    await driver.findElement(By.name("txtUsername")).sendKeys("Admin");
    await driver.findElement(By.name("txtPassword")).sendKeys("admin123");
    await await driver.findElement(By.id("btnLogin")).click();
    for(u = 0; u < data.length; u++){
        if(data[u].Is_Created == "Yes"){
            continue;
        }
        else{
            await driver.get('https://opensource-demo.orangehrmlive.com/index.php/pim/addEmployee');
            await driver.findElement(By.name("chkLogin")).click();
            await driver.findElement(By.name("firstName")).sendKeys(data[u].First_Name);
            await driver.findElement(By.name("lastName")).sendKeys(data[u].Last_Name);
            await driver.findElement(By.name("user_name")).sendKeys(data[u].Username);
            await driver.findElement(By.name("user_password")).sendKeys(data[u].Password);
            await driver.findElement(By.name("re_password")).sendKeys(data[u].Password);
            await driver.findElement(By.name("employeeId")).clear();
            await driver.findElement(By.name("employeeId")).sendKeys(data[u].Employee_id);
            await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div/div[2]/form/fieldset/p/input")).click();
            await new Promise(r => setTimeout(r, 2000));
            await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div/div[2]/div[2]/form/fieldset/p/input")).click();
            
            if(data[u].Gender == "Male"){
                await driver.findElement(By.id("personal_optGender_1")).click();
            }
            else{
                await driver.findElement(By.id("personal_optGender_2")).click();
            }
            await new Promise(r => setTimeout(r, 250));
            await driver.findElement(By.id("personal_cmbMarital")).sendKeys(data[u].Marital_Status);
            await new Promise(r => setTimeout(r, 250));
            await driver.findElement(By.id("personal_cmbNation")).sendKeys(data[u].Nationality);
            await new Promise(r => setTimeout(r, 250));
            //console.log(data[u].Date_of_Birth);
            await driver.findElement(By.name("personal[DOB]")).clear();
            await driver.findElement(By.name("personal[DOB]")).sendKeys(data[u].Date_of_Birth);
            await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div/div[2]/div[2]/form/fieldset/p/input")).click();
            await new Promise(r => setTimeout(r, 750));
            var url = await driver.getCurrentUrl();
            //console.log(url);
            var partOfUrl = url.substring(url.indexOf("Number")+7,url.length);
            //console.log(partOfUrl);
            await new Promise(r => setTimeout(r, 250));
            await driver.get('https://opensource-demo.orangehrmlive.com/index.php/pim/contactDetails/empNumber/' + partOfUrl);
            await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div/div[2]/div[2]/form/fieldset/p/input")).click();

            await driver.findElement(By.id("contact_street1")).sendKeys(data[u].Address_Street_1);
            await driver.findElement(By.id("contact_street2")).sendKeys(data[u].Address_Street_2);
            await driver.findElement(By.id("contact_city")).sendKeys(data[u].City);
            await driver.findElement(By.id("contact_province")).sendKeys(data[u].State_Province);
            await driver.findElement(By.id("contact_emp_zipcode")).sendKeys(data[u].Zip_Postal_Code);
            await driver.findElement(By.id("contact_country")).sendKeys(data[u].Country);
            await driver.findElement(By.id("contact_emp_hm_telephone")).sendKeys(data[u].Home_Telephone);
            await driver.findElement(By.id("contact_emp_mobile")).sendKeys(data[u].Mobile);
            await driver.findElement(By.id("contact_emp_work_telephone")).sendKeys(data[u].Work_Telephone);
            await driver.findElement(By.id("contact_emp_work_email")).sendKeys(data[u].Work_Email);
            await driver.findElement(By.id("contact_emp_oth_email")).sendKeys(data[u].Other_Email);

            await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div/div[2]/div[2]/form/fieldset/p/input")).click();
            await new Promise(r => setTimeout(r, 750));

            await driver.get('https://opensource-demo.orangehrmlive.com/index.php/pim/viewJobDetails/empNumber/' + partOfUrl);
            await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div[1]/div[2]/div[2]/form/fieldset/p/input[1]")).click();

            await driver.findElement(By.id("job_job_title")).sendKeys(data[u].Job_Title);
            await driver.findElement(By.id("job_emp_status")).sendKeys(data[u].Employment_Status);
            await driver.findElement(By.id("job_eeo_category")).sendKeys(data[u].Job_Category);
            await driver.findElement(By.id("job_joined_date")).clear();
            await driver.findElement(By.id("job_joined_date")).sendKeys(data[u].Joined_Date);
            await driver.findElement(By.id("job_sub_unit")).sendKeys(data[u].Sub_Unit);
            await driver.findElement(By.id("job_contract_start_date")).clear();
            await driver.findElement(By.id("job_contract_start_date")).sendKeys(data[u].Start_Date);
            await driver.findElement(By.id("job_contract_end_date")).clear();
            await driver.findElement(By.id("job_contract_end_date")).sendKeys(data[u].End_Date);
            await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div[1]/div[2]/div[2]/form/fieldset/p/input[1]")).click();
            data[u].Is_Created = "Yes";
            await new Promise(r => setTimeout(r, 750));
            
        }

    }
    //var today = new Date();
    wb.Sheets["Sheet1"] = xlsx.utils.json_to_sheet(data);

    xlsx.writeFile(wb,"EmployeeData.xls",{cellDates:true});

    for(u = 0; u < data.length; u++){
        if(data[u].Leave_From != " " && data[u].Leave_Assigned == "No"){
            await driver.get('https://opensource-demo.orangehrmlive.com/index.php/leave/assignLeave');
            await driver.findElement(By.name('assignleave[txtEmployee][empName]')).sendKeys(data[u].First_Name + " " + data[u].Last_Name);
            await driver.findElement(By.id("assignleave_txtLeaveType")).sendKeys(data[u].Leave_Type);
            await driver.findElement(By.id("assignleave_txtFromDate")).clear();
            await driver.findElement(By.id("assignleave_txtFromDate")).sendKeys(data[u].Leave_From);
            await new Promise(r => setTimeout(r, 750));
            await driver.findElement(By.id("assignleave_txtToDate")).clear();
            await driver.findElement(By.id("assignleave_txtToDate")).sendKeys(data[u].Leave_To);
            await new Promise(r => setTimeout(r, 750));
            //await driver.findElement(By.id("assignleave_duration_duration")).sendKeys(data[u].Duration);
            await driver.findElement(By.id("assignleave_txtComment")).sendKeys(data[u].Comment);
            //await driver.findElement(By.name('assignleave[txtEmployee][empName]')).sendKeys();
            //await driver.actions().mouseMove({x: 50, y: 0}).perform();
            
            //await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div[1]/div[2]/form/fieldset/p/input")).click();
            await new Promise(r => setTimeout(r, 750));
            var myButton = await driver.findElement(By.xpath("//*[@id='assignBtn']"));
            await new Promise(r => setTimeout(r, 750));
            
            myButton.getAttribute("value");
            await new Promise(r => setTimeout(r, 750));
            //await driver.actions().(myButton).click().perform();
            //await driver.findElement(By.xpath("/html/body/div[1]/div[3]/div[4]/div[3]/input[1]")).click();
            await new Promise(r => setTimeout(r, 750));
        }
    }
    
    wb.Sheets["Sheet1"] = xlsx.utils.json_to_sheet(data);

    xlsx.writeFile(wb,"EmployeeData.xls",{cellDates:true});

    //await driver.get('https://opensource-demo.orangehrmlive.com/index.php/pim/viewEmployeeList/reset/1');
    //for(u = 0; u < data.length; u++){
        
        //console.log(await driver.findElements(By.className("left")).find(item => item == data[u].Employee_id));
        
        //console.log(await (await driver.findElement(By.xpath("//*[@id='resultTable']/tbody/tr[4]/td[2]/a"))).getText());
    //}
    
    //wb.Sheets["Sheet1"] = xlsx.utils.json_to_sheet(data);

    //xlsx.writeFile(wb,"EmployeeData.xls",{cellDates:true});
}

login();

