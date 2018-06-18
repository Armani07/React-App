const Application = require('spectron').Application;
const path = require('path');

export default class spectronHelper {

    initialiseSpectron() {
        let electronPath = path.join(__dirname, "../../node_modules", ".bin", "electron");
        const appPath = path.join(__dirname, '..', 'main.js');
        if (process.platform === "win32") {
            electronPath += ".cmd";
        }
        return new Application({
            path: electronPath,
            args: [appPath],

            startTimeout: 20000
        });
    }
}