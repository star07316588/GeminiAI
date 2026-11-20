using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using MES.Models.Dto.CPAutoMonitor;
using MES.Repositories;

namespace MES.Services
{
    public class CPAutoMonitorService
    {
        private readonly CPAutoMonitorRepository _repository;

        // 狀態常數 (請依實際 DB 字串確認)
        private const string STATUS_INIT = "Init";
        private const string STATUS_AUTO_CHECKFAIL = "AutoCheckFail";
        private const string STATUS_AUTO_ERROR = "AutoError";
        private const string STATUS_AUTO_EXEC = "AutoExec";
        private const string STATUS_AUTO_FAIL = "AutoFail";
        private const string STATUS_MANUAL_EXEC = "ManualExec";
        private const string STATUS_AUTO_OK = "AutoOK";
        private const string STATUS_MANUAL_OK = "ManualOK";

        // 排序優先度常數
        private const int LEVEL_ERROR = 1;
        private const int LEVEL_INIT_TIMEOUT = 2;
        private const int LEVEL_EXEC_TIMEOUT = 3;
        private const int LEVEL_MOVEOUT_TIMEOUT = 4;
        private const int LEVEL_SORTRUN_NO_CPAUTO = 5;
        private const int LEVEL_AUTO_EXEC = 98;
        private const int LEVEL_MANUAL_EXEC = 99;

        public CPAutoMonitorService(CPAutoMonitorRepository repository)
        {
            _repository = repository;
        }

        public async Task<IEnumerable<CPAutoMonitorItemDto>> ExecuteMonitorQueryAsync()
        {
            var results = new List<CPAutoMonitorItemDto>();
            DateTime dbNow = await _repository.GetDbSysDateAsync();

            // ── 第一部分：TBL_CP_AUTO_COMPLETE 狀態檢查 ──
            var activeStatuses = new[]
            {
                STATUS_INIT, STATUS_AUTO_CHECKFAIL, STATUS_AUTO_ERROR,
                STATUS_AUTO_EXEC, STATUS_AUTO_FAIL, STATUS_MANUAL_EXEC
            };

            var part1Rows = await _repository.GetCpAutoActiveRecordsAsync(activeStatuses);
            foreach (var row in part1Rows)
            {
                string ruleName = !string.IsNullOrWhiteSpace(row.LotCompleteStatus) ? "LotComplete" : "LotEnd";
                string mainStatus = !string.IsNullOrWhiteSpace(row.LotCompleteStatus) ? row.LotCompleteStatus : row.LotEndStatus;

                string action = null;
                int priority = 0;
                string tdsSumTime = string.Empty;

                if (mainStatus == STATUS_AUTO_FAIL || mainStatus == STATUS_AUTO_CHECKFAIL || mainStatus == STATUS_AUTO_ERROR)
                {
                    priority = LEVEL_ERROR;
                    action = $"Auto {ruleName} Fail, 請手動執行{ruleName}.\nPlease execute {ruleName}.";
                }
                else if (mainStatus == STATUS_INIT)
                {
                    var (isReady, maxUpdateTime) = await CheckTdsTestSumAsync(row);
                    if (isReady && IsOverLimit(maxUpdateTime, dbNow, TimeSpan.FromMinutes(2)))
                    {
                        priority = LEVEL_INIT_TIMEOUT;
                        tdsSumTime = maxUpdateTime;
                        action = $"資料已到齊但是尚未Auto {ruleName}, 請手動執行{ruleName}.\nPlease execute {ruleName}.";
                    }
                }
                else if (mainStatus == STATUS_AUTO_EXEC)
                {
                    if (IsOverLimit(row.UpdateTime, dbNow, TimeSpan.FromMinutes(10)))
                    {
                        priority = LEVEL_EXEC_TIMEOUT;
                        action = $"Auto {ruleName} 執行過久, 請手動介入.\nPlease execute {ruleName}.";
                    }
                    else
                    {
                        priority = LEVEL_AUTO_EXEC;
                        action = $"Auto {ruleName} 執行中, 請稍候.\nKeep Monitor.";
                    }
                }
                else if (mainStatus == STATUS_MANUAL_EXEC)
                {
                    priority = LEVEL_MANUAL_EXEC;
                    action = $"Manual {ruleName} 執行中, 請注意是否執行過久.\nKeep Monitor.";
                }

                if (priority > 0)
                {
                    results.Add(new CPAutoMonitorItemDto
                    {
                        Priority = priority,
                        IsAlert = IsAlertPriority(priority),
                        EqID = row.EqID,
                        LotID = row.LotID,
                        StepNo = row.StepNo,
                        StepName = row.StepName,
                        LotStatus = row.LotStatus,
                        CurrentRule = row.CurrentRule,
                        TdsTestSumTime = tdsSumTime,
                        CPAutoStatusTime = !string.IsNullOrWhiteSpace(row.UpdateTime) ? row.UpdateTime : row.CreateTime,
                        LotEndStatus = row.LotEndStatus,
                        LotCompleteStatus = row.LotCompleteStatus,
                        Comments = row.Comments,
                        Action = action
                    });
                }
            }

            // ── 第二部分：MoveOut 超時檢查 (5 分鐘) ──
            var part2Rows = await _repository.GetMoveOutCandidatesAsync(new[] { STATUS_AUTO_OK, STATUS_MANUAL_OK });
            foreach (var row in part2Rows)
            {
                if (IsOverLimit(row.UpdateTime, dbNow, TimeSpan.FromMinutes(5)))
                {
                    results.Add(new CPAutoMonitorItemDto
                    {
                        Priority = LEVEL_MOVEOUT_TIMEOUT,
                        IsAlert = true,
                        EqID = row.EqID,
                        LotID = row.LotID,
                        StepNo = row.StepNo,
                        StepName = row.StepName,
                        LotStatus = row.LotStatus,
                        CurrentRule = row.CurrentRule,
                        CPAutoStatusTime = row.UpdateTime,
                        LotEndStatus = row.LotEndStatus,
                        LotCompleteStatus = row.LotCompleteStatus,
                        Comments = row.Comments,
                        Action = "MoveOut超時，請確認CommMgm系統重工記錄.\n若排除上述異常，請手動執行強制MoveOut或Reposition Rule.\nPlease execute move out forcibly after confirm CommMgm system \"rework\" record were done."
                    });
                }
            }

            // ── 第三部分：SortRun 自動化紀錄初始化失敗查詢 ──
            var part3Rows = await _repository.GetSortRunLotsAsync();
            foreach (var row in part3Rows)
            {
                bool hasRecord = await _repository.HasSortRunCpAutoRecordAsync(row.LotID);
                if (!hasRecord)
                {
                    results.Add(new CPAutoMonitorItemDto
                    {
                        Priority = LEVEL_SORTRUN_NO_CPAUTO,
                        IsAlert = true,
                        EqID = row.EqID,
                        LotID = row.LotID,
                        StepNo = row.StepNo,
                        StepName = row.StepName,
                        LotStatus = row.LotStatus,
                        CurrentRule = row.CurrentRule,
                        Action = "沒有對應的SortRun自動化紀錄，請重新LotStart或手動執行後續作業.\nPlease execute LotEnd and LotComplete."
                    });
                }
            }

            // 依 Priority 升冪排序
            return results.OrderBy(x => x.Priority).ToList();
        }

        // 對應原 VB6 CheckTdsTestSum
        private async Task<(bool IsReady, string MaxUpdateTime)> CheckTdsTestSumAsync(CpAutoRawDto lot)
        {
            string parentLot = GetParentLot(lot.LotID);
            var activeWafers = (await _repository.GetActiveWaferIdsAsync(lot.LotID)).ToList();
            if (!activeWafers.Any()) return (false, string.Empty);

            var tdsRows = (await _repository.GetTdsTestSumAsync(parentLot, lot.StepNo, lot.TimeHereSince)).ToList();
            if (!tdsRows.Any()) return (false, string.Empty);

            string maxUpdateTime = string.Empty;

            foreach (var waferId in activeWafers)
            {
                var matchedRows = tdsRows.Where(x => string.Equals(x.WaferId?.Trim(), waferId?.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
                if (!matchedRows.Any()) return (false, string.Empty);

                foreach (var r in matchedRows)
                {
                    bool hasTdsQty = !string.IsNullOrWhiteSpace(r.TdsGoodDie);
                    bool hasReviseQty = !string.IsNullOrWhiteSpace(r.ReviseGoodDie);
                    bool hasReason = !string.IsNullOrWhiteSpace(r.ReasonCode);

                    if ((!hasTdsQty && !hasReviseQty) || (hasReviseQty && !hasReason))
                    {
                        return (false, string.Empty);
                    }

                    string updateTime = !string.IsNullOrWhiteSpace(r.UpdateTime) ? r.UpdateTime : r.CreateTime;
                    if (string.CompareOrdinal(updateTime ?? string.Empty, maxUpdateTime) > 0)
                    {
                        maxUpdateTime = updateTime;
                    }
                }
            }

            return (true, maxUpdateTime);
        }

        private string GetParentLot(string lotId)
        {
            if (string.IsNullOrEmpty(lotId)) return string.Empty;
            int dotIdx = lotId.IndexOf('.');
            string baseLot = dotIdx > 0 ? lotId.Substring(0, dotIdx) : lotId;
            return baseLot.Length > 8 ? baseLot.Substring(0, 8) : baseLot;
        }

        // 解析 MES 18碼時間格式 (YYYYMMDD HH24MISS000) 並比對是否超時
        private bool IsOverLimit(string mesTimeStr, DateTime dbNow, TimeSpan limit)
        {
            if (string.IsNullOrWhiteSpace(mesTimeStr) || mesTimeStr.Length < 15) return false;
            if (DateTime.TryParseExact(mesTimeStr.Substring(0, 15), "yyyyMMdd HHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedTime))
            {
                return dbNow > parsedTime.Add(limit);
            }
            return false;
        }

        private bool IsAlertPriority(int priority)
        {
            return priority == LEVEL_ERROR ||
                   priority == LEVEL_INIT_TIMEOUT ||
                   priority == LEVEL_EXEC_TIMEOUT ||
                   priority == LEVEL_MOVEOUT_TIMEOUT ||
                   priority == LEVEL_SORTRUN_NO_CPAUTO;
        }
    }
}
