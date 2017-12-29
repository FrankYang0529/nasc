const fs = require('fs-promise');
const request = require('request-promise');
const puppeteer = require('puppeteer');

const getSatelliteURL = async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  await page.goto('http://www.cwb.gov.tw/V7/observe/satellite/Sat_T.htm?type=0');
  await page.click('a.satelliteImg[onclick="SelectArea(s3p,\'1\')"]');
  const imgSrc = await page.$eval('#im', el => el.src);
  await browser.close();

  return imgSrc;
};

const getRaderURL = async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  await page.goto('http://www.cwb.gov.tw/V7/observe/radar/?type=1');
  const imgSrc = await page.evaluate(() => {
    return document.body.querySelector('#viewer2 > img').src;
  })
  await browser.close();

  return imgSrc;
};

const downloadImage = async (imgSrc, fileName) => {
  const imgBody = await request.get(imgSrc, { encoding : null });
  await fs.writeFile(fileName, imgBody);
}

const downloadPlaneAlert = async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  await page.goto('https://aiss.anws.gov.tw/aes/ext/airspaceNotam2.jsp?userid=18d8c82a7d748d');
  await page.waitFor(3000);
  const img = await page.$('#map');
  await img.screenshot({ 'path': 'plane_alert.png' });
  await browser.close();
}

const getWeatherMetar = async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  const res = await page.goto('https://aiss.anws.gov.tw/aes/AwsClientMetar?stations=RCKH,RCBS,RCDC,RCFN,RCLY,RCNN,RCQC');
  const resStr = await res.text();
  const weatherMetar = resStr.replace(/\\\\\\/g, '').replace(/  /g, ' ').replace(/  /g, ' ');
  await fs.writeFile('weather_metar.json', weatherMetar, 'utf8');
  await browser.close();
}

(async () => {
  const satelliteURL = await getSatelliteURL();
  const raderURL = await getRaderURL();
  await downloadImage(satelliteURL, 'sateelite.jpg');
  await downloadImage(raderURL,'rader.jpg');
  await downloadPlaneAlert();
  await getWeatherMetar();
})();
