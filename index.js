const XLSX =require('xlsx');
const puppeteer = require('puppeteer');
const cheerio = require('cheerio');
const clc = require("cli-color");


let today = new Date().getDay(); // get the current day (0 for Sunday, 1 for Monday, etc.)
let sheet; //Select the sheet according current day

switch (today) {
    case 0:
        sheet = "Sunday";
        break;
    case 1:
        sheet = "Monday";
        break;
    case 2:
        sheet = "Tuesday";
        break;
    case 3:
        sheet = "Wednesday";
        break;
    case 4:
        sheet = "Thursday";
        break;
    case 5:
        sheet = "Friday";
        break;
    case 6:
        sheet = "Saturday";
        break;
}

// Now we can access the data in the sheet for the current day
// We can use the SheetJS library to parse the Excel file and access the data in the sheet

// Read the Excel file
let workbook = XLSX.readFile("Excel.xlsx");
let worksheet = workbook.Sheets[sheet];
let data = XLSX.utils.sheet_to_json(worksheet);

async function run(){
    // Open a browser and create a new page 
    const browser = await puppeteer.launch({ headless: false}); 
    const page = await browser.newPage(); // Create a new Page 
    
    // Loop through the keywords and write them in the search bar one at a time 
    for (let i= 0; i< data.length; i++){
        const keyword = data[i].__EMPTY_1;

        // go to google and write keyword to google search bar 
        await page.goto('https://www.google.com'); 
        await page.focus('input[name="q"]');
        await page.keyboard.type(keyword); 

        await page.waitForSelector('[role="listbox"]');
        await new Promise(resolve => setTimeout(resolve, 1000)); // wait 1 second
        
        // wait for the suggestions to load the HTML
        const html = await page.evaluate(() => document.body.innerHTML);

        // Use Cheerio to parse the HTML and the suggested options 
        const $ =  cheerio.load(html);
        const options = $('[role="listbox"] [role="option"]').map((i, el) =>$(el).text()).get();
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Find the Longest and Shortest options
        options.sort((a, b) => a.length - b.length);
        await new Promise(resolve => setTimeout(resolve, 1000));
        const shortest = options.shift();
        const longest = options.pop();
        
        console.log(clc.blue(`Shortest Option for ${keyword} -> `,shortest));
        console.log(clc.green(`Longest  Option for ${keyword} -> `,longest));


        // write the data to the appropriate cells in the Excel file
        const row = i + 3 ; // the data start on row 3 
        worksheet[`D${row}`] = { v: longest };
        worksheet[`E${row}`] = { v: shortest };
        
        // write the update data to the Excel file
        XLSX.writeFile(workbook, './Excel.xlsx'); 
        await new Promise(resolve => setTimeout(resolve, 1000));
    }
    browser.close();
}
run()

