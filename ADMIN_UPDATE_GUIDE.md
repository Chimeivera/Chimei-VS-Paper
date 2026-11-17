# 論文資料管理者更新方式

正式網站：https://chimeivera.github.io/Chimei-VS-Paper/

## 權限
- 一般訪客只能看數據與匯出報表，沒有網頁資料匯入功能。
- 有此 GitHub repository 寫入權限的管理者，才能變更共用資料。
- 不使用公開的網頁管理密碼；**請不要把 GitHub 密碼或 Personal Access Token 放到網頁**。

## 更新（全站同步）
1. 先備份目前完整原始 Excel 與 GitHub 版本。
2. 用院內核對完畢的「完整」110–115 論文申請表，不是只含增量的檔案；保留原始「流水號、申請序號、申請者、論文類別、獎金、論文名稱」等欄位。
3. 登入擁有 repository 寫入權限的 GitHub 帳號，進入 `Chimeivera/Chimei-VS-Paper` 的 `data/` 資料夾，上傳檔案並命名為 `論文申請資料.xlsx`；同名時選擇取代，提交到 main。
4. 開啟 Actions → `Publish approved paper Excel data`，確認成功。程式會比對已發布資料，拒絕重複申請序號、歷史紀錄遺失或被改寫。
5. Actions 成功後，更新版會寫入網站 `index.html`，所有訪客重新開啟正式網址就會看到共用的新資料。GitHub Pages 部署可能有短暫延遲，請以網站數字核對。
6. 如 Actions 失敗，**舊網站資料保持不變**。不可為了通過檢查而直接刪改舊紀錄；先查出差異。

## 還原
網站檔案版本可由 GitHub commit 紀錄還原；管理者應同時保存每次發布前的 Excel。請不要直接對正式 branch 強制回退或刪除資料，先確認當次異動範圍。

## 資料保護
GitHub Pages 與公開 repository 不是院內權限系統。論文申請檔可能含人事編號與內部簽核資訊；在上傳原始 Excel 之前，請先確認醫院允許將該資料放到公開 repository。若不能公開，應改用受控的後端儲存與驗證機制。
