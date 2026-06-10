using MES.Net.Infrastructure.Logging;
using MES.Net.Infrastructure.Repository;
using MES.Net.Infrastructure.Repository.Query;
using MES.Net.Infrastructure.Repository.Table;
using MES.Net.Shared.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static MES.Net.Shared.Constants.GlobalConstants;
using static MES.Net.Shared.DTOs.WsMergeQuery;

namespace MES.Net.Application.Services.Query
{
    public interface IWsMergeQueryService
    {
        Task<IEnumerable<WsMergeQuery>> GetLotMergeInfoAsync(string lotId, bool isHistoryTab);
        Task<IEnumerable<SorterControlInfo>> GetSorterControlInfoAsync(string lotId, string action);
        Task ResendSapAsync(ResendSapRequest request);
    }

    public class WsMergeLotService : IWsMergeQueryService
    {
        private readonly IWsMergeQueryRepository _lotRepo;

        public WsMergeLotService()
        {
        }

        public WsMergeLotService(IWsMergeQueryRepository lotRepo)
        {
            _lotRepo = lotRepo;
        }

        public async Task<IEnumerable<WsMergeQuery>> GetLotMergeInfoAsync(string lotId, bool isHistoryTab)
        {
            // Step-1 : 先檢查是否為 Virtual Lot ?
            bool isVirtualLot = await _lotRepo.CheckVirtualLotExistsAsync(lotId);

            if (isVirtualLot)
            {
                // 查 Virtual Merge
                return await _lotRepo.GetVirtualMergeInfoAsync(lotId, isHistoryTab);
            }
            else
            {
                // 查一般 Lot
                bool isAllLot = await _lotRepo.CheckAllLotExistsAsync(lotId);

                if (!isAllLot)
                {
                    throw new ArgumentException($"Lot '{lotId}' is not exist in MES !! 不存在於 MES !!");
                }

                // 查測後併批 tables (Inkless Merge)
                return await _lotRepo.GetInklessMergeInfoAsync(lotId, isHistoryTab);
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
        public async Task ResendSapAsync(ResendSapRequest request)
        {
            // 將傳入的陣列交給資料層執行 Update
            await _lotRepo.UpdateSendSapFlagAsync(request.Items, request.UserId);
        }
    }
}
