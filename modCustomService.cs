using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using MES.Net.Shared.DTOs.Print;
using MES.Net.Infrastructure.Repository.Print;

namespace MES.Net.Application.Services.Print
{
    public class RunCardService : IRunCardService
    {
        private readonly IRunCardRepository _repository;

        public RunCardService(IRunCardRepository repository)
        {
            _repository = repository;
        }

        public async Task<RunCardResponse> GetRunCardDataAsync(PrintRunCardRequest request)
        {
            // ... (前面的基本資料與 Spec 撈取邏輯) ...

            var response = new RunCardResponse { LotId = request.LotId };
            
            // 💡 應用場景 1：取得 Lot Type 來決定 Future Action 的過濾條件
            string lotType = await _repository.GetLotTypeAsync(request.LotId);
            // 稍後在撈取 Future Actions (GetFutureActionsAsync) 時，可以把 lotType 傳進去當作參數

            // ... (撈取生產履歷) ...
            var histories = await _repository.GetStepHistoryAsync(request.LotId, "您的PlanId");
            
            // 💡 應用場景 2 & 3：走訪履歷時，動態還原當下的機台屬性與 Lot 屬性
            foreach (var history in histories)
            {
                if (history.TrackOutTime.HasValue)
                {
                    // 取得當時機台的 EqpId (這裡假設 FWTRACKOUT 表裡面有存，或另外關聯)
                    string eqpId = "機台ID"; 

                    // 還原當時的 SubSys 狀態
                    string subSys1 = await _repository.GetEqpAttrAsync(eqpId, "SubSys1", history.TrackOutTime.Value);
                    
                    // 還原當時 Lot 的 Recipe 狀態
                    string pgName = await _repository.GetLotAttrAsync(request.LotId, "PgName", history.TrackOutTime.Value);
                    
                    // 將還原的狀態塞入 response 中
                    history.Equipment = string.IsNullOrEmpty(subSys1) ? eqpId : $"{eqpId} ({subSys1})";
                    history.Recipe = pgName;
                }
            }
            
            // ... (略) ...

            return response;
        }
    }
}
