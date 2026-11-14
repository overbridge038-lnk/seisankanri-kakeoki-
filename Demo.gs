/**
 * 生産管理しくみ  デモ用ウェブアプリ（プレゼンで触ってもらう用） ver 2026-11-14a
 *
 * 本番のデータ置き場（Code.gs）とは、別の Apps Script プロジェクトに入れる。
 * 1. https://script.google.com → 新しいプロジェクト → このファイルの中身を貼って保存
 * 2. ファイル「＋」→ HTML → 名前を index にして、「デモ用ファイル」の中身を全部貼って保存
 *    （デモ用ファイル＝index.html の 設定 →「デモ用ファイルをつくる」でできるHTML）
 * 3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *      次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 * 4. できた URL（…/exec）を、プレゼンで見てもらう人に送る
 *
 * 本番のスプレッドシート・URL・合言葉には、つながらない（デモ用ファイルには入らない）。
 * 触った内容は、触った人の端末にだけ残る。
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('生産管理しくみ（デモ）')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1,viewport-fit=cover');
}
