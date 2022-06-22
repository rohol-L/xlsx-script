import xlsx_script from "../src/index.mjs"
import fs from 'fs'

function readFileAsync(fileName) {
    return new Promise((resolve, reject) => {
        fs.readFile(fileName, function (err, data) {
            if (err) reject(err);
            else resolve(data.toString())
        })
    })
}

async function test(fileName) {
    let [xs, fileData] = await Promise.all([
        xlsx_script.loadFile("./" + fileName + '.xlsx'),
        readFileAsync("./" + fileName + '.json')
    ])
    const data = JSON.parse(fileData)
    const startTime = new Date();
    //xs.logOutput = 2
    xs.render(data)
    console.info("render:" + fileName, (new Date() - startTime) / 1000)
    xs.save("./out_" + fileName + '.xlsx')
}

async function testAll() {
    await test("test1")
    await test("test2")
}

testAll()
