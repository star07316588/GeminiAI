using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MES.Models.Dto.MergeSplitQuery;
using MES.Repositories;

namespace MES.Services
{
    public class MergeSplitQueryService
    {
        private readonly MergeSplitQueryRepository _repository;

        public MergeSplitQueryService(MergeSplitQueryRepository repository)
        {
            _repository = repository;
        }

        public async Task<IEnumerable<MergeSplitQueryResultDto>> ExecuteQueryAsync(MergeSplitQueryRequest req)
        {
            var results = new List<MergeSplitQueryResultDto>();

            // 1. 取得 Merge / Split 的基礎資料
            if (req.Activity == "Merge" || req.Activity == "All")
            {
                var mergeData = await _repository.GetMergeRecordsAsync(req);
                results.AddRange(mergeData);
            }

            if (req.Activity == "Split" || req.Activity == "All")
            {
                var splitData = await _repository.GetSplitRecordsAsync(req);
                results.AddRange(splitData);
            }

            // 排序 (對應 order by PARENTLOT, TXNTIME)
            results = results.OrderBy(x => x.ParentLotId).ThenBy(x => x.TimeStamp).ToList();

            // 若筆數大於限制 (例如 gsQueryLimitCnt)，可以在此處拋出 Exception 或截斷
            if (results.Count > 1000)
            {
                throw new System.Exception("查詢筆數過多，請增加查詢條件。");
            }

            // 2. 針對每一筆紀錄，去撈出對應的 Qty (對應 VB6 中 Do While oRawSqlRecordSet.EOF 裡面的行為)
            foreach (var row in results)
            {
                // 對 StepId 進行字串截斷處理 (Left(STEPID, 5))
                row.StepId = !string.IsNullOrEmpty(row.StepId) && row.StepId.Length > 5 ? row.StepId.Substring(0, 5) : row.StepId;

                var wipTxns = await _repository.GetWipTransactionsAsync(row.ParentLotId, row.ChildLotId, row.SysId);

                foreach (var txn in wipTxns)
                {
                    if (txn.WipId == row.ParentLotId)
                    {
                        if (txn.Attribute == "ChipQty")
                        {
                            row.OriParCQty = txn.OriginalValue;
                            if (string.IsNullOrEmpty(row.AftParCQty)) row.AftParCQty = txn.NewValue;
                        }
                        else // WaferQty
                        {
                            row.OriParWQty = txn.OriginalValue;
                            if (string.IsNullOrEmpty(row.AftParWQty)) row.AftParWQty = txn.NewValue;
                        }
                    }
                    else
                    {
                        if (txn.Attribute == "ChipQty")
                        {
                            row.ChildCQty = (row.Activity == "Merge") ? txn.OriginalValue : txn.NewValue;
                        }
                        else
                        {
                            row.ChildWQty = (row.Activity == "Merge") ? txn.OriginalValue : txn.NewValue;
                        }
                    }
                }

                // 補償 Child Lot Qty 邏輯 (若未在上述 Query 找到)
                if (string.IsNullOrEmpty(row.ChildCQty) || string.IsNullOrEmpty(row.ChildWQty))
                {
                    var childHistory = await _repository.GetChildLotHistoryQtyAsync(row.ChildLotId, row.TimeStamp);
                    foreach (var hist in childHistory)
                    {
                        if (hist.Attribute == "ChipQty" && string.IsNullOrEmpty(row.ChildCQty))
                        {
                            row.ChildCQty = hist.NewValue;
                        }
                        else if (hist.Attribute == "WaferQty" && string.IsNullOrEmpty(row.ChildWQty))
                        {
                            row.ChildWQty = hist.NewValue;
                        }
                    }
                }
            }

            return results;
        }
    }
}
