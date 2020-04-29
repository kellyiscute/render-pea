#!/usr/bin/env node

const puppeteer = require('puppeteer');
const fs = require('fs');
const process = require('process');
const yargs = require('yargs');
const xlsx = require('xlsx');
const cheerio = require('cheerio');
const path = require('path')
const os = require('os')

//TODO: read config

const fuck = async () => {
    const isPkg = typeof process.pkg !== 'undefined'
    let chromiumExecutablePath
    if (os.platform() == 'win32') {
        console.log(puppeteer.executablePath())
        chromiumExecutablePath = 'chrome-win\\chrome.exe'
    } else {
        chromiumExecutablePath = (isPkg ?
            puppeteer.executablePath().replace(
                /^.*?\/node_modules\/puppeteer\/\.local-chromium/,
                path.join(path.dirname(process.execPath), 'chromium')
            ) :
            puppeteer.executablePath()
        )
    }

    const browser = await puppeteer.launch({
        headless: true,
        executablePath: chromiumExecutablePath
    })
    const page = await browser.newPage();

    await page.goto('file://' + process.cwd() + '/to_html.html', {
        waitUntil: "load"
    })
    await page.pdf({
        path: out,
        format: 'A4',
        margin: {
            top: "20px",
            left: "20px",
            right: "20px",
            bottom: "20px"
        }
    });
    await browser.close()

};

const argv = yargs
    .command('fuck')
    .option('config')

let border = 'solid 1px'
let color = 'black'
let out = 'out.pdf'

const shit = () => {
    let file;
    if (!argv.argv._[0]) {
        console.log('fucked');
        return
    } else {
        file = argv.argv._[0]
    }
    if (argv.argv.config) {
        console.log(argv.argv.config);
        config = fs.readFileSync(argv.argv.config)
        config = JSON.parse(config)
        if (config.border) {
            border = config.border
        }
        if (config.color) {
            color = config.color
        }
    }
    if (argv.argv.out) {
        out = argv.argv.out
    }

    var workbook = xlsx.readFile(file);
    json = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
        header: 1
    });
    html = fs.readFileSync('./index.html').toString()
    html = html.replace('//!!DATA!!', JSON.stringify(json))
    html = html.replace('!!border!!', border)
    html = html.replace('!!border!!', border)
    html = html.replace('!!color!!', color)
    fs.writeFileSync(process.cwd() + '/to_html.html', html)

    fuck(html)
}

shit()