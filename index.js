const path = require("path");
const fs = require("fs");
const xl = require("xlsx");
const prompt = require("prompt-sync")();
// Make it easy to log
const log = (msg) => console.log(msg);

const location = prompt("Where is the input file? (Try dragging and dropping onto this window) ");
log("");
log("Reading file: " + location);
log("");
const workbook = xl.readFile(location, {
    type: "binary"
});
workbook.SheetNames.forEach(name => {
    log("Parsing sheet " + name + "...");
    try {
        const json = xl.utils.sheet_to_json(workbook.Sheets[name]);
        log("Converted " + json.length + " records!");
        log("Writing to file...");
        try {
            fs.writeFileSync(path.join(__dirname, `${name}.Formatted.json`), JSON.stringify(json));
            log("Success!");   
        } catch (err) {
            log("Error occurred.");
            console.error(err);
        }
    } catch (e) {
        log("Error occurred.");
        console.error(e);
    }
});
log("")
log("Completed!");
process.exit();
