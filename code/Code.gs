/**
 * @license
 * このソフトウェアは、MITライセンスのもとで公開されています。
 * This software is released under the MIT License.
 *
 * Copyright (c) 2024 Masaaki Maeta
 *
 * 以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「本ソフトウェア」）の複製を取得するすべての人に対し、本ソフトウェアを無制限に扱うことを無償で許可します。これには、本ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、および本ソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
 *
 * 上記の著作権表示および本許諾表示を、本ソフトウェアのすべての複製または重要な部分に記載するものとします。
 *
 * 本ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、本ソフトウェアに起因または関連し、あるいは本ソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
 *
 * --- (English Original) ---
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

const PARENT_FOLDER_NAME = "録画くん保存フォルダ";
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SUBFOLDER_SHEET_NAME = "シート1";
const SUBFOLDER_CELL = "B1";
const HISTORY_SHEET_NAME = "履歴";
const HISTORY_HEADERS = ["ファイル名", "保存日時", "フォルダパス", "ファイルリンク"];

function doGet(e) {
  createHistorySheetIfNotExists();
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('動画録画アプリ (画質・カメラ切替対応)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getOrCreateFolderIdByName(folderName, parentFolder = DriveApp.getRootFolder()) {
  if (!folderName || typeof folderName !== 'string' || folderName.trim() === '') {
    throw new Error("有効なフォルダ名が指定されていません。");
  }
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next().getId();
  } else {
    const newFolder = parentFolder.createFolder(folderName);
    Logger.log(`フォルダ "${folderName}" を作成しました。`);
    return newFolder.getId();
  }
}

function getSubFolderNameFromSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SUBFOLDER_SHEET_NAME);
    if (!sheet) return null;
    const subFolderName = sheet.getRange(SUBFOLDER_CELL).getValue().toString().trim();
    return subFolderName ? subFolderName.replace(/[\\\/:\*\?"<>\|]/g, '_') : null;
  } catch (e) {
    Logger.log(`サブフォルダ名の取得エラー: ${e.toString()}`);
    return null;
  }
}

function createHistorySheetIfNotExists() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.appendRow(HISTORY_HEADERS);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setFontWeight("bold");
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 300);
  }
  return sheet;
}

function addRecordToHistorySheet(fileName, folderPathText, folderUrl, fileUrl) {
  try {
    const sheet = createHistorySheetIfNotExists();
    const timestamp = new Date();
    const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    const folderLinkFormula = `=HYPERLINK("${folderUrl}","${folderPathText}")`;
    const fileLinkFormula = `=HYPERLINK("${fileUrl}","${fileName}")`;
    sheet.appendRow([fileName, formattedTimestamp, folderLinkFormula, fileLinkFormula]);
  } catch (e) {
    Logger.log(`履歴シートへの記録エラー: ${e.toString()}`);
  }
}

/**
 * Base64エンコードされた動画データをデコードし、WebMファイルとして保存します。
 * @param {string} videoDataUrl - "data:video/webm;base64,..." の形式の動画データURL。
 * @param {string} baseFileName - 保存するファイルのベース名 (拡張子なし)。
 * @return {Object} 保存処理の結果。
 */
function saveVideoFile(videoDataUrl, baseFileName) {
  try {
    if (!videoDataUrl || !baseFileName) {
      throw new Error("動画データまたはファイル名が無効です。");
    }
    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    let targetFolder, folderPathText, targetFolderUrl;
    const subFolderNameRaw = getSubFolderNameFromSheet();
    if (subFolderNameRaw) {
      const subFolderId = getOrCreateFolderIdByName(subFolderNameRaw, parentFolder);
      targetFolder = DriveApp.getFolderById(subFolderId);
      folderPathText = `${parentFolder.getName()} > ${targetFolder.getName()}`;
      targetFolderUrl = targetFolder.getUrl();
    } else {
      targetFolder = parentFolder;
      folderPathText = parentFolder.getName();
      targetFolderUrl = parentFolder.getUrl();
    }
    const parts = videoDataUrl.match(/^data:(.+?);base64,(.+)$/);
    if (!parts) throw new Error("無効なData URL形式です。");
    const mimeType = parts[1];
    const base64Data = parts[2];
    
    const finalFileName = `${baseFileName}.webm`;

    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, finalFileName);
    const file = targetFolder.createFile(blob);
    addRecordToHistorySheet(finalFileName, folderPathText, targetFolderUrl, file.getUrl());
    return {
      success: true,
      message: `ファイル "${finalFileName}" をドライブのフォルダ「${folderPathText}」に保存しました。`
    };
  } catch (error) {
    Logger.log(`saveVideoFileでエラーが発生しました: ${error.toString()}`);
    return {
      success: false,
      message: `ファイルの保存中にサーバー側でエラーが発生しました。`
    };
  }
}
