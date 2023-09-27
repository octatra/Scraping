const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

(async () => {
  const browser = await puppeteer.launch({ 
    headless: false,
    args:[
      '--start-maximized'
    ],
    defaultViewport: null,
   }); // Launch a non-headless browser for debugging
  const page = await browser.newPage();

  // Navigate to the login page
  await page.goto('https://lmsbebras.tigaharmoni.com/admin');

  // Fill in login credentials and submit the form
  await page.type('input[name="email"]', 'admin@gmail.com'); // Replace with your username
  await page.type('input[name="password"]', 'admin123'); // Replace with your password
  await page.click('button[type="submit"]');

 
  // console.log("Lewat");
  await page.waitForTimeout(2000);

  await page.waitForSelector('#sidebar-wrapper > ul > li:nth-child(7) > a');

  await page.click('#sidebar-wrapper > ul > li:nth-child(7) > a');
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');
  
  let column = 1;
  
   for (let index = 2; index <= 402; index++) {
    await page.waitForTimeout(2000);

   await page.waitForSelector(`#app > div > div.main-content > section > div.row > div > div > div.card-body.p-0 > div > table > tbody > tr:nth-child(${index}) > td:nth-child(5) > a`);
    // await page.waitForSelector(`#app > div > div.main-content > section > div.row > div > div > div.card-body.p-0 > div > table > tbody > tr:nth-child(${index}) > td:nth-child(4) > a`);
  
    // Mengklik elemen dengan selector yang Anda berikan
    await page.click(`#app > div > div.main-content > section > div.row > div > div > div.card-body.p-0 > div > table > tbody > tr:nth-child(${index}) > td:nth-child(5) > a`);
    await page.waitForTimeout(2000);

    const dataSiswa = await page.$eval(`#app > div > div.main-content > section > div.row > div > div > div.card-body > h5`, element => element.textContent);

    worksheet.getCell(1, column).value = dataSiswa;

    const data = await page.$eval(`#app > div > div.main-content > section > div.row > div > div > div.card-body > div > table > tbody > tr:nth-child(2) > td:nth-child(1) > h5`, element => element.textContent);
    const data2 = await page.$eval(`#app > div > div.main-content > section > div.row > div > div > div.card-body > div > table > tbody > tr:nth-child(2) > td:nth-child(2) > h5`, element => element.textContent);
    const cleanedData = data.replace(/\s+/g, '').replace(/\|/g, '\n');
    const cleanedData2 = data2.replace(/\s+/g, '').replace(/\|/g, '\n');

 // Menambahkan label "LCM" ke worksheet
 worksheet.getCell(2, column).value = "LCM";

 // Menambahkan data yang sudah dibersihkan ke dalam worksheet
 const dataArray = cleanedData.split('\n');
 dataArray.forEach((dataItem, rowIndex) => {
   worksheet.getCell(rowIndex + 3, column).value = dataItem;
 });

 // Menambahkan label "FYS" ke worksheet
 worksheet.getCell(2, column + 1).value = "FYS";

 // Menambahkan data FYS yang sudah dibersihkan ke dalam worksheet
 const dataArray2 = cleanedData2.split('\n');
 dataArray2.forEach((dataItem2, rowIndex) => {
   worksheet.getCell(rowIndex + 3, column + 1).value = dataItem2;
 });

 column += 3; // Melanjutkan ke kolom berikutnya

    await page.waitForTimeout(2000);
    await page.goBack();
  }

  // Simpan workbook sebagai file Excel
  await workbook.xlsx.writeFile('Data_Korelasi.xlsx');

  console.log('Data telah disimpan dalam file Excel (data.xlsx).');


  await browser.close();
})();
