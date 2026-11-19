using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MES.Models.Dto.MonitorPCQuery;
using MES.Repositories;

namespace MES.Services
{
    public class MonitorPCQueryService
    {
        private readonly MonitorPCQueryRepository _repository;

        public MonitorPCQueryService(MonitorPCQueryRepository repository)
        {
            _repository = repository;
        }

        public Task<MonitorPCInitDataDto> GetInitDataAsync()
        {
            return _repository.GetInitDataAsync();
        }

        public async Task<IEnumerable<MonitorPCQueryResultDto>> ExecuteQueryAsync(MonitorPCQueryRequest req)
        {
            var pcList = (await _repository.GetMonitorPCListAsync(req)).ToList();
            if (!pcList.Any()) return pcList;

            // 依 SystemName 快取最新版本清單，避免迴圈重複查詢 DB
            var maxVersionCache = new Dictionary<string, List<MonitorPCDetailDto>>(StringComparer.OrdinalIgnoreCase);

            foreach (var pc in pcList)
            {
                if (!maxVersionCache.TryGetValue(pc.SystemName, out var maxVersions))
                {
                    maxVersions = (await _repository.GetLatestVersionBySystemAsync(pc.SystemName)).ToList();
                    maxVersionCache[pc.SystemName] = maxVersions;
                }

                // 撈取該台電腦的有效檔案版本，轉成不分大小寫的 Dictionary
                var pcFiles = await _repository.GetVersionDetailByIpAsync(pc.IP, pc.SystemName, onlyActive: true);
                var pcFileMap = pcFiles
                    .GroupBy(x => x.FileName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.First().FileUpdateTime ?? string.Empty, StringComparer.OrdinalIgnoreCase);

                string versionFlag = "N";

                // 對應原 VB6 邏輯：當該電腦有抓到檔案時才判斷，只要任一檔案舊於最新版本即為 "N" 並中斷
                foreach (var maxVer in maxVersions)
                {
                    if (pcFileMap.TryGetValue(maxVer.FileName ?? string.Empty, out var currentUpdateTime))
                    {
                        versionFlag = "N";
                        if (string.CompareOrdinal(currentUpdateTime, maxVer.FileUpdateTime ?? string.Empty) >= 0)
                        {
                            versionFlag = "Y";
                        }
                        else
                        {
                            versionFlag = "N";
                            break;
                        }
                    }
                }

                pc.LastVersionFlag = versionFlag;
            }

            return pcList;
        }

        public async Task<IEnumerable<MonitorPCDetailDto>> GetDetailAsync(MonitorPCDetailRequest req)
        {
            if (string.Equals(req.IP, "ALL", StringComparison.OrdinalIgnoreCase))
            {
                return await _repository.GetLatestVersionBySystemAsync(req.SystemName);
            }

            return await _repository.GetVersionDetailByIpAsync(req.IP, req.SystemName, onlyActive: false);
        }
    }
}
