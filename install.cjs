// このアドインをExcelに登録する（レジストリへのsideload登録）。
// taskpane本体はGitHub Pagesでホストしているため、ローカルサーバーは不要。
// 同僚にこのフォルダごと渡せば、各自のPCで node install.cjs（またはinstall.bat）を
// 実行するだけで使えるようになる（Node.jsはこの登録コマンドのためだけに必要）。
// 使い方: node install.cjs
const { execFile } = require("child_process");
const path = require("path");

const KEY = "HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer";
const NAME = "dfc3fd23-ff19-4d33-b312-a15d117dd27d";
const MANIFEST_PATH = path.join(__dirname, "manifest.xml");

execFile("reg", ["add", KEY, "/v", NAME, "/t", "REG_SZ", "/d", MANIFEST_PATH, "/f"], (err, stdout, stderr) => {
  if(err){
    console.log("❌ レジストリ登録に失敗しました:", stderr.trim() || err.message);
    process.exitCode = 1;
    return;
  }
  console.log("✅ レジストリに登録しました。");
  console.log("");
  console.log("次にすること: Excelを起動し、ホームタブの「工程表ツール」ボタンから開く。");
  console.log("（表示されない場合は、Excelを完全に終了してから開き直してください）");
});
