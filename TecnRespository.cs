        // ==========================================
        // 取得單一 Level 的 TECN PGM 設定 (考量 Grade 溫度欄位)
        // ==========================================
        public async Task<dynamic> GetTecnPgmByLevelAsync(
            string eqType2, string runRule, string tempColumn, string testMode, 
            int level, string description, string tecnNo, string pgmNameCondition, 
            string path = null, string stepId = null) // 🌟 補上可選參數
        {
            // 組裝 Base SQL (對應 VB6 的 sSQL_1)
            string sql = $@"
                SELECT TECNNO, TECNLEVEL, PGNAME, PGID, STATUS, SOURCE, 
                       {tempColumn} as Temp,
                       CASE WHEN TO_CHAR(SYSDATE, 'YYYYMMDD HH24MISS') || '000' BETWEEN STARTTIME AND ENDTIME THEN 'N' ELSE 'Y' END as OverTime
                FROM TBL_TECN_PGM 
                WHERE TESTERTYPE = :EqType2 ";

            if (runRule == "Y")
            {
                sql += " AND REFRENCE = '引用' AND STATUS <> '失效已Confirm' ";
            }
            else
            {
                sql += @" AND STARTTIME <= TO_CHAR(SYSDATE,'YYYYMMDD HH24MISS') || '000' 
                          AND TO_CHAR(SYSDATE,'YYYYMMDD HH24MISS') || '000' <= ENDTIME 
                          AND ( (REFRENCE='引用' AND STATUS ='生效已Confirm') OR SOURCE ='Old' ) ";
            }

            // 組裝 Level 條件 (對應 VB6 的 sSQL_2)
        sql += " AND TESTMODE = :TestMode AND TECNLEVEL = :Level ";

            // 🌟 處理 TQAE 的專屬過濾條件
            if (!string.IsNullOrEmpty(path))
            {
                sql += " AND PATH = :Path ";
            }
            if (!string.IsNullOrEmpty(stepId))
            {
                sql += " AND STEPID = :StepId ";
            }

            // ... (後面的 SQL 組裝與 QueryFirstOrDefaultAsync 邏輯) ...
            
            // 記得要把 Path 與 StepId 放進 Dapper 參數裡
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new 
            { 
                EqType2 = eqType2, TestMode = testMode, Level = level, 
                Description = description, TecnNo = tecnNo, PgmNameCondition = pgmNameCondition,
                Path = path, StepId = stepId
            });
        }
        
        // ==========================================
        // 檢查 Lot 控制清單是否被標記 Delete
        // ==========================================
        public async Task<bool> IsLotTecnDeletedAsync(string tecnNo, string lotId)
        {
            string sql = @"
                SELECT ACTION FROM TBL_TBL_LOT_TECN_CONTROL_LIST 
                WHERE TECNNO = :TecnNo AND LOTID = :LotId AND DELETEFLAG = 'N'
                ORDER BY CREATETIME DESC, ACTIONTIME DESC
                FETCH FIRST 1 ROWS ONLY"; // Oracle 12c+ 語法; 若為舊版請用 ROWNUM = 1 子查詢
                
            string action = await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { TecnNo = tecnNo, LotId = lotId });
            return action == "Delete";
        }
