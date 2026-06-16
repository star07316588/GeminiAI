using System;
using System.Data;
using System.Linq;
using Dapper;

public class PerformanceRepository
{
    public int CloseCheck(
        IDbConnection con, string sType, string loginTitle,
        string year, string month, string stationId, 
        string title, string deptId, string shiftId)
    {
        // 1. 清理字串：移除括號與單引號 (讓 Dapper 參數化時不會有重複引號的問題)，並切成陣列
        // 使用 StringSplitOptions.RemoveEmptyEntries 避免切出空字串
        string[] stationArray = stationId.Replace("(", "").Replace(")", "").Replace("'", "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        string[] titleArray = title.Replace("(", "").Replace(")", "").Replace("'", "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        string[] shiftArray = shiftId.Replace("(", "").Replace(")", "").Replace("'", "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

        // 2. 判斷資料表名稱 (Table Name 無法參數化，後續用字串內插)
        string sTable = sType == "SUBJECT" 
            ? "dlowner.rbl_dl_performance_subject" 
            : "dlowner.rbl_dl_performance_manager";

        try
        {
            // --- 查詢 1：取得 Detail Items ---
            string sqlItems = $@"
                SELECT DISTINCT a.item, a.detailitem
                FROM rbl_dl_detailitem a, rbl_dl_item b
                WHERE a.station_id = b.station_id
                  AND a.title = b.title
                  AND a.item = b.item
                  AND a.DELETEFLAG = 'N'
                  AND b.DELETEFLAG = 'N'
                  AND b.subjective = 'Y'
                  AND a.station_id IN :StationIds 
                  AND a.title IN :TitleIds
                  AND (INSTR(input, :LoginTitle) <> 0 OR input = 'ALL')
            ";

            var items = con.Query(sqlItems, new { 
                StationIds = stationArray, 
                TitleIds = titleArray, 
                LoginTitle = loginTitle 
            }).ToList();

            // --- 進入核心檢核邏輯 ---
            foreach (var row in items)
            {
                foreach (string st in stationArray)
                {
                    foreach (string ti in titleArray)
                    {
                        foreach (string sh in shiftArray)
                        {
                            // 查詢 2：計算人數 (使用 ExecuteScalar 直接取回 int)
                            string sqlPeoples = @"
                                SELECT COUNT(*) 
                                FROM sbl_emp b 
                                WHERE b.title = :Title 
                                  AND b.station_id = :Station 
                                  AND b.shift_id = :Shift 
                                  AND b.dept_id = :DeptId 
                                  AND b.DLT_USER IS NULL
                            ";
                            int peoples = con.ExecuteScalar<int>(sqlPeoples, new { Title = ti, Station = st, Shift = sh, DeptId = deptId });

                            // 只有當人數 > 0 時，才需要去查績效表有無紀錄，節省 DB 效能
                            if (peoples > 0)
                            {
                                // 查詢 3：計算績效數量
                                string sqlCount = $@"
                                    SELECT COUNT(*) 
                                    FROM {sTable} a, sbl_emp b 
                                    WHERE a.emp_id = b.emp_id   
                                      AND a.month = :Month 
                                      AND a.year = :Year 
                                      AND a.item = :Item 
                                      AND a.detailitem = :DetailItem 
                                      AND b.title = :Title 
                                      AND b.station_id = :Station 
                                      AND b.shift_id = :Shift 
                                      AND b.dept_id = :DeptId 
                                ";

                                // 動態附加條件
                                if (sType == "MANAGER")
                                {
                                    sqlCount += " AND a.final IS NOT NULL ";
                                }

                                int cnt = con.ExecuteScalar<int>(sqlCount, new {
                                    Month = month,
                                    Year = year,
                                    Item = (string)row.item,
                                    DetailItem = (string)row.detailitem,
                                    Title = ti,
                                    Station = st,
                                    Shift = sh,
                                    DeptId = deptId
                                });

                                // 若人數 > 0 但卻沒有績效紀錄 (cnt == 0)，代表檢核失敗
                                if (cnt == 0)
                                {
                                    // 直接 return 0！取代原本 Java 冗長的 bbreak=true 和層層 break
                                    return 0; 
                                }
                            }
                        }
                    }
                }
            }

            // 若所有迴圈跑完都沒有觸發 return 0，代表檢核全部通過
            return 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            return 0; // 發生例外時，比照原本 Java 邏輯回傳 0
        }
    }
}
