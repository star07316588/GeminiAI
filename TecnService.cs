public async Task<TecnPgmRecipeAttrDto> GetTecnPgmRecipeAttrAsync(
            string ipnTecnNo, string lotId, string ipn, string path, string stepId, string stepName, 
            string eqType2, string prodGroupKey, string prodCode, string grade, string runRule)
        {
            var result = new TecnPgmRecipeAttrDto { RefOverTime = "N" };

            // 1. 決定溫度欄位 (Grade 轉換)
            string tempColumn = "TEMP_" + (string.IsNullOrEmpty(grade) ? "C" : grade.ToUpper());

            // 2. 決定 TestMode
            string testMode = stepName;
            bool isTqae = stepName.StartsWith("TQAE");
            if (stepName.StartsWith("SORT"))
            {
                testMode = "S" + stepName.Substring(4);
            }
            else if (isTqae)
            {
                // TODO: 呼叫 Get_TQAE_Mapping_Act_PGmode 取得 QCTestMode
                // 此處假設已經取得 mappedTestMode
                // testMode = mappedTestMode; 
            }

            // ==========================================
            // PASS 1: 尋找 PGM (配方)
            // ==========================================
            for (int iLevel = 1; iLevel <= 6; iLevel++)
            {
                if (iLevel == 5) continue;

                string description = GetDescriptionForLevel(iLevel, lotId, ipn, prodGroupKey, prodCode);
                
                // 若為 TQAE 站，VB6 邏輯會先嘗試原 stepName，找不到再嘗試 testMode (QC Test Mode)
                var record = await _repo.GetTecnPgmByLevelAsync(eqType2, runRule, tempColumn, stepName, iLevel, description, ipnTecnNo, null);
                
                if (record == null && isTqae)
                {
                    record = await _repo.GetTecnPgmByLevelAsync(eqType2, runRule, tempColumn, testMode, iLevel, description, ipnTecnNo, null);
                }
                else if (record == null && !isTqae)
                {
                    record = await _repo.GetTecnPgmByLevelAsync(eqType2, runRule, tempColumn, testMode, iLevel, description, ipnTecnNo, null);
                }

                if (record != null)
                {
                    // TODO: 若需要，可在此處實作 ComparePgmAttr 的驗證邏輯

                    string tecnNo = record.TECNNO?.ToString();
                    
                    // Lot 階被標記 Delete 的攔截
                    if (iLevel == 1 && await _repo.IsLotTecnDeletedAsync(tecnNo, lotId))
                    {
                        continue;
                    }

                    result.RefPgmTecnNo = tecnNo;
                    result.RefPgId = record.PGID?.ToString();
                    result.RefPgmSource = record.SOURCE?.ToString();
                    result.RefLevel = record.TECNLEVEL?.ToString();
                    
                    string rawPgmName = record.PGNAME?.ToString() ?? "";
                    
                    // TQAE AdvanGroup 字串置換邏輯 (需依據 bAdvanGroup 判斷，此處示意)
                    bool bAdvanGroup = false; // 需由前方邏輯取得
                    if (isTqae && bAdvanGroup)
                    {
                        if (rawPgmName.StartsWith("TF")) result.RefPgm = "TQ" + rawPgmName.Substring(2);
                        else if (rawPgmName.StartsWith("F")) result.RefPgm = "Q" + rawPgmName.Substring(1);
                        else result.RefPgm = rawPgmName;
                    }
                    else
                    {
                        result.RefPgm = rawPgmName;
                    }

                    // 記錄溫度 (如果有)
                    string tempValue = record.TEMP?.ToString();
                    if (!string.IsNullOrEmpty(tempValue))
                    {
                        result.RefTemp = tempValue;
                        result.RefTempTecnNo = tecnNo;
                        result.RefTempSource = record.SOURCE?.ToString();
                    }

                    string overTime = record.OVERTIME?.ToString();
                    if (result.RefOverTime != "Y") result.RefOverTime = overTime;

                    break; // 找到 PGM 就可以跳出 PASS 1 迴圈
                }
            }

            // ==========================================
            // PASS 2: 若 PGM 找到但沒包溫度，單獨尋找 TEMP
            // ==========================================
            if (!string.IsNullOrEmpty(result.RefPgm) && string.IsNullOrEmpty(result.RefTemp))
            {
                for (int iLevel = 1; iLevel <= 6; iLevel++)
                {
                    if (iLevel == 5) continue;
                    
                    string description = GetDescriptionForLevel(iLevel, lotId, ipn, prodGroupKey, prodCode);
                    
                    var record = await _repo.GetTecnPgmByLevelAsync(eqType2, runRule, tempColumn, testMode, iLevel, description, ipnTecnNo, result.RefPgm);
                    
                    if (record == null && isTqae)
                    {
                        record = await _repo.GetTecnPgmByLevelAsync(eqType2, runRule, tempColumn, testMode, iLevel, description, ipnTecnNo, result.RefPgm);
                    }

                    if (record != null)
                    {
                        result.RefTemp = record.TEMP?.ToString();
                        result.RefTempTecnNo = record.TECNNO?.ToString();
                        result.RefTempSource = record.SOURCE?.ToString();
                        
                        if (result.RefOverTime != "Y") result.RefOverTime = record.OVERTIME?.ToString();
                        if (string.IsNullOrEmpty(result.RefLevel)) result.RefLevel = record.TECNLEVEL?.ToString();
                        
                        break;
                    }
                }
            }

            return result;
        }

        // ==========================================
        // 輔助方法：根據 Level 取得 Description
        // ==========================================
        private string GetDescriptionForLevel(int level, string lotId, string ipn, string prodGroupKey, string prodCode)
        {
            switch (level)
            {
                case 1: return lotId;
                case 2:
                case 3: return ipn;
                case 4: return prodGroupKey;
                case 6: return prodCode;
                default: return "";
            }
        }
