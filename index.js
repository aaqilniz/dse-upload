const puppeteer = require('puppeteer');
const xlsx = require('node-xlsx').default;
require('dotenv').config()
const { USERNAME_DSE, PASSWORD_DSE } = process.env;
const motherTongue = 'S';
const medium = 'S';

async function uploadData() {
    try {
        // Launch the browser
        console.log('Launching the browser');
        const browser = await puppeteer.launch({
            headless: false, // Set to true for background execution
            defaultViewport: null,
            args: ['--start-maximized'] // Open browser maximized
        });

        // Create a new page
        const page = await browser.newPage();
        console.log('Launched the browser and opened a new page.');
        // Navigate to the login page
        await page.goto('https://dseshyd.gos.pk/Login.aspx?ReturnUrl=%2fschool%2fdashboard.aspx', {
            waitUntil: 'networkidle0'
        });
        console.log('Navigated to the dse site.');
        // Wait for the username input field
        await page.waitForSelector('#HeadLoginView_logId_UserName');

        // Enter username 
        await page.type('#HeadLoginView_logId_UserName', USERNAME_DSE);

        // Enter password
        await page.type('#HeadLoginView_logId_Password', PASSWORD_DSE);

        // Click login button
        await page.click('#HeadLoginView_logId_LoginButton');

        // Wait for navigation to complete
        await page.waitForNavigation({ waitUntil: 'networkidle0' });
        console.log('Successfully logged into the site.');

        await page.goto('https://dseshyd.gos.pk/enrol/default.aspx', { waitUntil: 'networkidle0' });
        console.log('Navigated to the enrollment page .');
        await page.click('#ContentPlaceHolder1_grvSchools_btnSelect_0');
        console.log('Loading enrollment form.');

        console.log('processing the excel data.');
        const allClassesGrView = xlsx.parse(
            `${__dirname}/all-classes-gr-view.xlsx`,
            { cellDates: true }
        );
        console.log('processed excel data.');
        let counter = 1;
        const allRecords = allClassesGrView[1]['data'];
        for (let i = 2; i < allRecords.length; i++) {
            const record = allRecords[i];
            let grNumber = record[1];
            const studentName = record[2];
            const fatherName = record[3];
            let religion = record[4];
            const caste = record[5];
            const placeOfBirth = record[6];
            let dateOfBirth = record[7];
            const previousSchool = record[11];
            let dateOfAdmission = record[12];
            let admissionClass = record[13];
            let currentClass = record[14];
            const classSection = record[15] || 'Green';
            const remarks = record[16];
            let studentNadraId = record[17] || '0000-00000000-0';
            let parentCnic = record[18] || '0000-00000000-0';
            let parentCellNo = record[19] || '0000-0000000';
            let gender = record[22];
            studentNadraId = studentNadraId.toString();
            parentCnic = parentCnic.toString();
            parentCellNo = parentCellNo.toString();
            if (!remarks || remarks !== 'New Admission') {
                continue; // no need to do the entry!
            }

            if (dateOfBirth) {
                dateOfBirth = new Date(dateOfBirth);
                // Extract the date parts (year, month, day) to keep it in local time
                const year = dateOfBirth.getFullYear();
                const month = String(dateOfBirth.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
                const day = String(dateOfBirth.getDate()).padStart(2, '0');
                dateOfBirth = `${year}-${month}-${day}`;
            }
            if (dateOfAdmission) {
                // Extract the date parts (year, month, day) to keep it in local time
                const year = dateOfAdmission.getFullYear();
                const month = String(dateOfAdmission.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
                const day = String(dateOfAdmission.getDate()).padStart(2, '0');
                dateOfAdmission = `${year}-${month}-${day}`;
            }
            console.log(`\n----------Record # ${counter}----------`);
            counter++;
            if (
                !grNumber ||
                !studentName ||
                !fatherName ||
                !caste ||
                !religion ||
                !placeOfBirth ||
                !dateOfBirth ||
                !previousSchool ||
                !dateOfAdmission ||
                !admissionClass ||
                !currentClass ||
                !classSection ||
                !studentNadraId ||
                !parentCnic ||
                !parentCellNo ||
                !gender
            ) {
                console.log(`failed to create record with GR# ${grNumber} Name ${studentName} Father's Name ${fatherName}`);
                continue
            }
            grNumber = record[1].toString();
            religion = mapReligion(religion);
            admissionClass = mapClass(record[13]);
            currentClass = mapClass(record[14]);
            gender = mapGender(record[22]);
            await delay();
            await page.evaluate((selector) => {
                if (document.querySelector(selector)) {
                    document.querySelector(selector).value = '';
                }
            }, '#ContentPlaceHolder1_txtGRNo');
            console.log('Verifying if the entry already exists');
            await page.type('#ContentPlaceHolder1_txtGRNo', grNumber);
            await page.click('#ContentPlaceHolder1_btnCheckGrNo');
            await delay(1000 * 10);
            let element = await page.$('#ContentPlaceHolder1_lblMsg');
            let value = await page.evaluate(el => el.textContent, element);
            if (value === 'This G.R No already exists.') {
                console.log(`This G.R No ${grNumber} already exists. This entry will be skipped.`);
                continue;
            } else if (value === '') {
                console.log('No existing record found. Proceeding with new entry.');
            }

            // await delay(1000 * 30);
            console.log('*****creating new entry.*****');
            await page.type('#ContentPlaceHolder1_RequiredFieldValidator7', dateOfAdmission);
            await page.select('#ContentPlaceHolder1_drdClass', admissionClass);
            await page.select('#ContentPlaceHolder1_drdCurrentClass', currentClass);
            await page.select('#ContentPlaceHolder1_drdSection', classSection);
            await page.select('#ContentPlaceHolder1_drdMotherTongue', motherTongue);
            await page.select('#ContentPlaceHolder1_drdMedium', medium);
            await page.select('#ContentPlaceHolder1_drdGender', gender);
            await page.select('#ContentPlaceHolder1_drdReligion', religion.toUpperCase());
            await page.type('#ContentPlaceHolder1_txtName', studentName);
            await page.type('#ContentPlaceHolder1_txtFName', fatherName);
            await page.type('#ContentPlaceHolder1_txtStudentCaste', caste);
            await page.type('#ContentPlaceHolder1_txtDateOfBirth', dateOfBirth);
            await page.type('#ContentPlaceHolder1_txtPlaceOfBirth', placeOfBirth);
            await page.type('#ContentPlaceHolder1_txtPreviousSchool', previousSchool);
            await page.type('#ContentPlaceHolder1_txtParentCNICNo', parentCnic);
            await page.type('#ContentPlaceHolder1_txtParentsCellNo', parentCellNo);
            await page.type('#ContentPlaceHolder1_txtNADRAID', studentNadraId);
            // await delay(1000 * 60 * 0.5);

            await page.click('#ContentPlaceHolder1_btnCreateStudentID'); // saving the entry!
            await delay(1000);
            element = await page.$('#ContentPlaceHolder1_lblMsg');
            value = await page.evaluate(el => el.textContent, element);
            console.log('[[[[[[value]]]]]]');
            console.log(value);
            console.log('[[[[[[value]]]]]]');
            if (value === 'Record saved successfully...!') {
                console.log(`successfully created the entry with GR# ${grNumber}`);
            } else {
                console.log(`failed created the entry with GR# ${grNumber}`);

            }
            // Wait for navigation to complete
            continue;
            // process.exit(0);
            // // districtOfBirth // options  BDN DDU HYD JSR MTR SJL TAR TMK TTA KCC KCE KCK KCM KCS KCW JBD KMR LRK QST SKP MPK THR UKT NFZ SBA SGR GTK KPR SKR Other
            // const districtOfBirth = 'HYD';
            // await page.select('#ContentPlaceHolder1_drdDistrictOfBirth', districtOfBirth);

            // await delay(1000 * 30);

            // // talukaOfBirth // options  Hyderabad: HDCT HDRT LTFT QSDT
            // // ContentPlaceHolder1_drdTalukaOfBirth
            // const talukaOfBirth = 'HDRT';
            // await page.select('#ContentPlaceHolder1_drdDistrictOfBirth', talukaOfBirth.toUpperCase());

        }
        console.log('finsihed uploading data.');
        process.exit(0);
    } catch (error) {
        console.error(error);
    }
}

// Run the login function
uploadData();


const mapReligion = (religion) => {
    switch (religion.toUpperCase()) {
        case 'ISLAM':
            return 'MUSLIM'
        case 'HINDU':
            return 'HINDU'
    }
};
const mapClass = (studentClass) => {
    switch (studentClass) {
        case 'VI':
            return '6'
        case 'VII':
            return '7'
        case 'VIII':
            return '8'
        case 'IX':
            return '9'
        case 'X':
            return '10'
    }
}


const mapGender = (gender) => {
    switch (gender) {
        case 'B':
            return 'BOY'
        case 'G':
            return 'GIRL'
    }
}

const delay = async (duration) => {
    return new Promise((resolve) => {
        setTimeout(() => {
            resolve();
        }, duration || 10000);
    })
}