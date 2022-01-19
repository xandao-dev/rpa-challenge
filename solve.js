const puppeteer = require('puppeteer');
const xlsx = require('node-xlsx');

const RPA_CHALLENGE_URL = 'http://www.rpachallenge.com/';
const CHALLENGE_XLSX_FILE = `${__dirname}/challenge.xlsx`;
const IS_HEADLESS = true;
const REPEAT_TIMES = 10;

const rpaSelectors = {
	start: 'body > app-root > div > app-rpa1 > div > div > div > button',
	name: "[ng-reflect-name='labelFirstName']",
	lastName: "[ng-reflect-name='labelLastName']",
	company: "[ng-reflect-name='labelCompanyName']",
	role: "[ng-reflect-name='labelRole']",
	address: "[ng-reflect-name='labelAddress']",
	email: "[ng-reflect-name='labelEmail']",
	phone: "[ng-reflect-name='labelPhone']",
	submitButton: 'body > app-root > div > app-rpa1 > div > div > form > input',
	result: '.message2',
};

const getUserData = () => {
	const worksheet = xlsx.parse(CHALLENGE_XLSX_FILE, {
		header: ['name', 'lastName', 'company', 'role', 'address', 'email', 'phone'],
		blankrows: false,
	});
	const usersData = worksheet[0].data.slice(1);
	return usersData;
};

const configPages = async (browser) => {
	const pages = await browser.pages();
	const pageRPA = pages[0];

	await pageRPA.setDefaultNavigationTimeout(0);
	await pageRPA.setRequestInterception(true);

	pageRPA.on('request', (req) => {
		if (req.resourceType() === 'image' || req.resourceType() === 'font' || req.resourceType() === 'stylesheet') {
			req.abort();
		} else {
			req.continue();
		}
	});
	return { pageRPA };
};

const openRpaChallenge = async (pageRPA) => await pageRPA.goto(RPA_CHALLENGE_URL, { waitUntil: 'load' });

const start = async (pageRPA) => {
	await pageRPA.waitForSelector(rpaSelectors['start']);
	await pageRPA.click(rpaSelectors['start']);
};

const fillUserData = async (pageRPA, userData) => {
	await pageRPA.waitForSelector(rpaSelectors['name']);
	let name = await pageRPA.evaluate(`document.querySelector("${rpaSelectors['name']}").value`);
	if (!name) {
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['name']}").value = "${userData.name}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['lastName']}").value = "${userData.lastName}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['company']}").value = "${userData.company}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['role']}").value = "${userData.role}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['address']}").value = "${userData.address}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['email']}").value = "${userData.email}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['phone']}").value = "${userData.phone}"`);
	} else {
		while (name) {
			name = await pageRPA.evaluate(`document.querySelector("${rpaSelectors['name']}").value`);
		}
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['name']}").value = "${userData.name}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['lastName']}").value = "${userData.lastName}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['company']}").value = "${userData.company}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['role']}").value = "${userData.role}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['address']}").value = "${userData.address}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['email']}").value = "${userData.email}"`);
		await pageRPA.evaluate(`document.querySelector("${rpaSelectors['phone']}").value = "${userData.phone}"`);
	}
};

const goToNextPage = async (pageRPA) => {
	await pageRPA.waitForSelector(rpaSelectors['submitButton']);
	await pageRPA.click(rpaSelectors['submitButton']);
};

const getResults = async (pageRPA) => {
	await pageRPA.waitForSelector(rpaSelectors['result']);
	const result = await pageRPA.evaluate(`document.querySelector("${rpaSelectors['result']}").innerText`);
	console.log(`Results: ${result}`);
	return parseFloat(result.split('in ').pop().split(' ')[0]);
};

(async () => {
	const usersData = getUserData();
	const browser = await puppeteer.launch({ headless: IS_HEADLESS, devtools: false });
	const { pageRPA } = await configPages(browser);

	let averageTime = 0;
	for (let repeat = 0; repeat < REPEAT_TIMES; repeat++) {
		await openRpaChallenge(pageRPA);
		await start(pageRPA);
		for (let userData of usersData) {
			await fillUserData(pageRPA, userData);
			await goToNextPage(pageRPA);
		}
		averageTime += await getResults(pageRPA);
	}
	console.log(`Average time: ${averageTime / REPEAT_TIMES} milliseconds`);
	process.exit(0);
})();
