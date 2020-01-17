const puppeteer = require('puppeteer');
const excelToJson = require('convert-excel-to-json');
let prompt = require('prompt');


const config = {
    loadingOptions:{
        headless: true,
    },
    await:600,
};

let meterList = [];

let schema = {
    properties: {
      login: {
        pattern: /^[a-zA-Z\s\-]+$/,
        message: 'Name must be only letters',
        required: true,
      },
      password: {
        hidden: true,
        replace: '*'
      }
    }
  };
 
  prompt.start();

  prompt.get(schema, function (err, result) {
      main(result.login,result.password);
  });
///////////////////////////////////////////////////////////////////////////// login module

main = (login,password)=>{
  const result = excelToJson({
    sourceFile: '../listExcel/list.xlsx',
});

result.Sheet1.forEach((item)=>{
  meterList.push(item.A.toString());
});

const percent = (A , B) =>{
  let result = (B*100)/A;
  console.log(`${result}%`)
}

const time = (string , start , stop)=>{
  var currentDate = new Date();
  console.log(`${string} ${currentDate.getHours()} : ${currentDate.getMinutes()} : ${currentDate.getSeconds() < 10 ?  `0${currentDate.getSeconds()}` : currentDate.getSeconds()}`);
};
/////////////////////////////////////////////////////////////////////////// Excel sheet to json
(async () => {
  
  time('Start time: ');
    
  const browser = await puppeteer.launch(config.loadingOptions);
  const page = await browser.newPage();
  await page.goto('http://ukpbrvs045:50000/manufacturing/com/sap/me/activity/client/ActivityManager.jsp');
  await page.type('#logonuidfield',login);
  await page.type('#logonpassfield', password);
  await page.click('.urBtnStdNew');
  await page.goto('http://ukpbrvs045:50000/manufacturing/com/sap/me/wpmf/client/template.jsf?WORKSTATION=OPERATION_DEF&ACTIVITY_ID=DEF_OPER_POD');

  await page.waitFor(1000);
  const Op = await page.$('input[value="OP_REPAIRS"]');
  await Op.click({clickCount:3});
  await page.waitFor(config.await);
  await Op.type('OP_HIGH_VOLTAGE');
  await page.waitFor(config.await);

  const Res = await page.$('input[value="DEFAULT"]')
  await Res.click({clickCount:3});
  await page.waitFor(config.await);
  await Res.type('R_MAN_HV_001');
  await page.waitFor(config.await);

  
  for(let i = 0; i<meterList.length; i++){
    const Sfc = await page.$('input[maxlength="128"]');
    await Sfc.click({clickCount:3});
    await page.waitFor(config.await);
    await Sfc.type(meterList[i]);
    await page.waitFor(config.await);

    const Start = await page.$('span[title="Start"]');
    await Start.click();
    await page.waitFor(config.await);

    const Complete = await page.$('span[title="Complete"]');
    await Complete.click();
    await page.waitFor(config.await);
    await page.waitFor(600);
    percent(meterList.length,i+1);
  }

  await page.screenshot({path: 'example.png'});
  await browser.close();
  time('Finish time:');
})().catch(e => console.error("Wrong Loggin or Password. CTR + C to close program"));
}
 
