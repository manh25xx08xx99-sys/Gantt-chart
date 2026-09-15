// このアドインをExcelに登録する（レジストリへのsideload登録 + ログイン時にサーバーを
// 自動起動するショートカットの作成）。1人1台・自分のPCで完結する設定なので、
// 同僚にこのフォルダごと渡せば、各自のPCで node install.cjs（またはinstall.bat）を
// 実行するだけで使えるようになる。
// 使い方: node install.cjs
const { execFile } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");

const KEY = "HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer";
const NAME = "dfc3fd23-ff19-4d33-b312-a15d117dd27d";
const HERE = __dirname;
const MANIFEST_PATH = path.join(HERE, "manifest.xml");
const STARTUP_DIR = path.join(
  os.homedir(),
  "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup"
);
const STARTUP_BAT = path.join(STARTUP_DIR, "m5-gantt-server-autostart.bat");

function registerSideload(){
  return new Promise((resolve, reject) => {
    execFile("reg", ["add", KEY, "/v", NAME, "/t", "REG_SZ", "/d", MANIFEST_PATH, "/f"], (err, stdout, stderr) => {
      if(err) reject(new Error(stderr.trim() || err.message));
      else resolve();
    });
  });
}

function createAutoStartShortcut(){
  const script = [
    "@echo off",
    'cd /d "' + HERE + '"',
    'start "" /min node server.js',
  ].join("\r\n") + "\r\n";
  fs.mkdirSync(STARTUP_DIR, { recursive: true });
  fs.writeFileSync(STARTUP_BAT, script, "utf8");
}

(async () => {
  try{
    await registerSideload();
    console.log("✅ レジストリに登録しました（Excelの「工程表ツール」ボタンから開けます）。");
  }catch(err){
    console.log("❌ レジストリ登録に失敗しました:", err.message);
    process.exitCode = 1;
    return;
  }

  try{
    createAutoStartShortcut();
    console.log("✅ Windowsログイン時にサーバーが自動起動するように設定しました。");
    console.log("   （" + STARTUP_BAT + "）");
  }catch(err){
    console.log("⚠ 自動起動の設定に失敗しました（手動でstart-server.batを実行してください）:", err.message);
  }

  console.log("");
  console.log("次にすること:");
  console.log("1. install.bat（またはstart-server.bat）でサーバーを今すぐ起動する");
  console.log("2. Excelを起動し、ホームタブの「工程表ツール」ボタンから開く");
})();
