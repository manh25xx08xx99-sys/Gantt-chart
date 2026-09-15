// Excel からこのアドインを取り除く（レジストリの sideload 登録を削除する）
// 使い方: node uninstall.cjs
const { execFile } = require("child_process");
const KEY = "HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer";
const NAME = "dfc3fd23-ff19-4d33-b312-a15d117dd27d";

execFile("reg", ["delete", KEY, "/v", NAME, "/f"], (err, stdout, stderr) => {
  if(err){
    console.log("削除できませんでした（既に削除済みの可能性があります）:", stderr.trim() || err.message);
  }else{
    console.log("レジストリから登録を削除しました。Excel を再起動すると表示されなくなります。");
  }
});
