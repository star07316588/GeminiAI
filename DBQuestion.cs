        public async Task UpdateSendSapFlagAsync(IEnumerable<ResendSapItem> items, string userId)
        {
            // 產生更新時間 (對應原本的 to_char(sysdate, 'yyyymmddhh24miss') || '00')
            string updateTime = DateTime.Now.ToString("yyyyMMdd HHmmss") + "000";

            if (_dbConnection.State != ConnectionState.Open)
            {
                _dbConnection.Open();
            }

            // 開啟資料庫交易
            var transaction = _dbConnection.BeginTransaction();

            try
            {
                foreach (var item in items)
                {
                    string sql = $@"
                        UPDATE {item.UpdateTable}
                        SET SENDSAPFLAG = 'N',
                            SENDSAPTIME = '',
                            UPDATEUSERID = :UserId,
                            UPDATETIME = :UpdateTime
                        WHERE ROWID = :RowId 
                          AND NVL(SENDSAPFLAG, 'N') <> 'N'";

                    // 執行 SQL 並綁定參數與交易
                    await _dbConnection.ExecuteAsync(sql, new
                    {
                        UserId = userId,
                        UpdateTime = updateTime,
                        RowId = item.RowId
                    }, transaction);
                }

                // 若所有迴圈內的更新都沒發生例外，則提交交易
                transaction.Commit();
            }
            catch
            {
                // 若發生任何錯誤，復原所有已更新的資料
                transaction.Rollback();
                throw; // 將例外往上拋給 Controller 處理
            }
        }


        public async Task<IEnumerable<SorterControlInfo>> GetSorterControlInfoAsync(string lotId, string action)
        {
            // 第一次查詢：查詢尚未完成的 Sorter 作業
            var data = await _lotRepo.GetActiveSorterControlsAsync(lotId, action);

            // 若第一段 SQL (colRS) 沒有資料
            if (data == null || !data.Any())
            {
                // 第二次查詢 (colRS_2)：檢查是否其實已經完成了？
                bool hasFinished = await _lotRepo.CheckAnySorterControlExistsAsync(lotId, action);

                if (hasFinished)
                {
                    // 對應 MsgBox: Has finish Sorter operation
                    throw new InvalidOperationException($"Lot '{lotId}' has finish Sorter '{action}' operation !! 已完成 Sorter '{action}' 作業 !!");
                }
                else
                {
                    // 對應 MsgBox: is no data found
                    throw new InvalidOperationException($"Lot '{lotId}' is no '{action}' data found !! 查無 '{action}' 資料 !!");
                }
            }

            return data;
        }
