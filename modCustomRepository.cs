using Dapper;
using System;
using System.Data;
using System.Collections.Generic;
using System.Threading.Tasks;
using MES.Net.Shared.DTOs.Print;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace MES.Net.Infrastructure.Repository.Print
{
    public interface IRunCardRepository
    {
        // ... (保留原本的 QueryLotsAsync, GetLotSpecInfoAsync, 等方法)

        // 新增：共用輔助查詢方法
        Task<string> GetLotTypeAsync(string lotId);
        Task<string> GetStepPathAsync(string planId, string stepSeq);
        Task<string> GetLotAttrAsync(string lotId, string attrName, DateTime txnTime);
        Task<string> GetEqpAttrAsync(string eqpId, string attrName, DateTime txnTime);
    }

    public class RunCardRepository : IRunCardRepository
    {
        private readonly IDbConnection _dbConnection;
        // ... (保留建構子與原本的方法)

        /// <summary>
        /// 判斷批號型態 (Normal / Eng / Rework)
        /// </summary>
        public async Task<string> GetLotTypeAsync(string lotId)
        {
            string sql = @"
                SELECT ERUNTICNO, SAPRWNO 
                FROM TBL_LOT_INFO 
                WHERE LOT_ID = :LotId";
            
            // 使用 dynamic 承接，方便判斷欄位是否為 null
            var result = await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId });

            if (result == null) return "Normal";

            string erunTicNo = result.ERUNTICNO?.ToString();
            string sapRwNo = result.SAPRWNO?.ToString();

            // 仿照 VB 的判斷邏輯
            if (!string.IsNullOrEmpty(erunTicNo)) return "Eng";
            if (!string.IsNullOrEmpty(sapRwNo)) return "Rework";
            
            return "Normal";
        }

        /// <summary>
        /// 取得未來的途程 (Path)
        /// </summary>
        public async Task<string> GetStepPathAsync(string planId, string stepSeq)
        {
            // 將 VB6 拼湊的 SQL 轉換為標準的 ANSI JOIN 語法，提高可讀性與效能
            string sql = @"
                SELECT e.PATH 
                FROM fwprocessplan a
                INNER JOIN fwflatplan b ON b.planname = a.name AND b.planversion = a.activeversion
                INNER JOIN fwflatplan_n2m c ON b.sysid = c.fromid AND c.linkname = 'planNodes'
                INNER JOIN fwflatnode d ON c.toid = d.sysid
                INNER JOIN TBL_ROUTE_STEP_LIST e ON e.STEP_HANDLE = d.handle
                WHERE a.name = :PlanId 
                  AND b.RevState = 'Frozen' 
                  AND d.StepSeq = :StepSeq 
                  AND e.ROUTE_ID = :PlanId 
                  AND e.DELETE_FLAG = 'N'";

            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { PlanId = planId, StepSeq = stepSeq });
        }

        /// <summary>
        /// 取得批號在特定時間點的屬性值
        /// </summary>
        public async Task<string> GetLotAttrAsync(string lotId, string attrName, DateTime txnTime)
        {
            string sql = @"
                SELECT NEWVALUE 
                FROM FWWIPTRANSACTION 
                WHERE WIPID = :LotId 
                  AND ATTRIBUTE = :AttrName 
                  AND TXNTIMESTAMP = (
                      SELECT MAX(TXNTIMESTAMP) 
                      FROM FWWIPTRANSACTION 
                      WHERE WIPID = :LotId 
                        AND ATTRIBUTE = :AttrName 
                        AND TXNTIMESTAMP <= :TxnTime
                  )";
            // Dapper 會自動處理 DateTime 到 Oracle 日期的參數綁定
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { LotId = lotId, AttrName = attrName, TxnTime = txnTime });
        }

        /// <summary>
        /// 取得機台在特定時間點的屬性值
        /// </summary>
        public async Task<string> GetEqpAttrAsync(string eqpId, string attrName, DateTime txnTime)
        {
            string sql = @"
                SELECT ATTRIBUTEMAGNITUDE 
                FROM VIEW_FWEQPHISTORY 
                WHERE EQUIPMENTNAME = :EqpId 
                  AND EVENTNAME = 'FwEquipmentModifyAttributeTxn' 
                  AND ATTRIBUTENAME = :AttrName 
                  AND TIME = (
                      SELECT MAX(TIME) 
                      FROM FWEQPHISTORY 
                      WHERE EQUIPMENTNAME = :EqpId 
                        AND EVENTNAME = 'FwEquipmentModifyAttributeTxn' 
                        AND ATTRIBUTENAME = :AttrName 
                        AND TIME < :TxnTime
                  )";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { EqpId = eqpId, AttrName = attrName, TxnTime = txnTime });
        }
    }
}
