// Excel からこのアドインを取り除く（レジストリの sideload 登録と、
// ログイン時自動起動用のバッチファイルを削除する）
// 使い方: node uninstall.cjs
const { execFile } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");

const KEY = "HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer";
const NAME = "dfc3fd23-ff19-4d33-b312-a15d117dd27d";
const STARTUP_BAT = path.join(
  os.homedir(),
  "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup",
  "m5-gantt-server-autostart.bat"
);

execFile("reg", ["delete", KEY, "/v", NAME, "/f"], (err, stdout, stderr) => {
  if(err){
    console.log("削除できませんでした（既に削除済みの可能性があります）:", stderr.trim() || err.message);
  }else{
    console.log("レジストリから登録を削除しました。Excel を再起動すると表示されなくなります。");
  }
});

try{
  if(fs.existsSync(STARTUP_BAT)){
    fs.unlinkSync(STARTUP_BAT);
    console.log("ログイン時自動起動の設定も削除しました。");
  }
}catch(err){
  console.log("自動起動設定の削除に失敗しました（手動で削除してください）:", err.message);
}
