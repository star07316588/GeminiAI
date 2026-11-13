using System;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using log4net;

namespace MIFNet.FTP
{
    class EIFPCDMSMES04 : MIFFTP
    {
        /**
         * Update B/E MES Database for modify PCDMS data
         * 
         * If this method is multi-thread safe, you can remove the synchronized 
         * modifier to improve performance.
         */
        private string _strutInfo = "EIFPCDMSMES04";
        private string MifName = "H1B104";

        public EIFPCDMSMES04(string[] args)
            : base(args)
        {
            //sMifNameGET = MifName;
        }

        protected override void setSendData()
        {
            StringBuilder sSql = new StringBuilder();
            string sSystime = string.Empty;
            int intProcessLine = 0;

            PCDMSFunction _pcdms = new PCDMSFunction();
            MIFUtil _mesdb = new MIFUtil();
            MIFUtil _data = new MIFUtil();
            _mesdb.DBTool("MESDB");
            List<string> _allDataGET = new List<string>();


            // 1. 設定來源資料夾與搜尋條件
            string folderPath = @"D:\Report\PCDMS\H1B104_Listen";
            string searchPattern = "H1B104*";

            // 2. 設定歸檔資料夾路徑 (D:\mif\source_file\archive\YYYYMM)
            string archiveBasePath = @"D:\mif\source_file\archive";
            string archivePath = Path.Combine(archiveBasePath, DateTime.Now.ToString("yyyyMM"));

            // 3. 檢查來源資料夾是否存在
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                return;
            }

            // 4. 檢查並建立歸檔(Archive)資料夾
            if (!Directory.Exists(archivePath))
            {
                Directory.CreateDirectory(archivePath);
            }

            // 5. 取得檔案清單 (因為後續要搬移檔案，這裡必須先轉成確定的陣列，避免列舉中途檔案消失報錯)
            string[] files = Directory.GetFiles(folderPath, searchPattern);


            sSystime = DateTime.Now.ToString("yyyyMMddHHmmss");

            _log.Info("*** Start to operate the executeUpdate_" + MifName + " at " + sSystime);

            foreach (string file in files)
            {
                _allDataGET.AddRange(File.ReadLines(file));

                string fileName = Path.GetFileName(file);
                string destinationPath = Path.Combine(archivePath, fileName);

                sSystime = DateTime.Now.ToString("yyyyMMddHHmmss");

                _log.Info("*** Start to operate the executeUpdate_" + MifName + " at " + sSystime);


                if (_allDataGET.Count > 0)
                {
                    foreach (string strLine in _allDataGET)
                    {
                        sSql = Makeup_IntH1B104_SQL(strLine, _mesdb, _log);
                        _mesdb.ExecNonQuery(sSql.ToString());
                        sSql.Clear();
                        sSql = _pcdms.Makeup_Repair_PcdLifeTime_SQL(strLine, _mesdb, _log);
                        _mesdb.ExecNonQuery(sSql.ToString());
                        sSql.Clear();
                        sSql = _pcdms.Makeup_Update_Grade_SQL(strLine, "Repair", _mesdb, _log);
                        _mesdb.ExecNonQuery(sSql.ToString());
                        sSql.Clear();
                        // Update_PcdQuality PCD健康管理需求
                        sSql = _pcdms.Update_PcdQuality(strLine, "Repair", _mesdb, _log);
                        if (sSql.ToString() != "NA")
                        {
                            _mesdb.ExecNonQuery(sSql.ToString());
                        }
                        sSql.Clear();
                        intProcessLine = intProcessLine + 1;

                        _log.Info("Total Record:" + intProcessLine);

                        _log.Info("Update executeUpdate_" + _strutInfo + " finished!");
                    }
                }

                // 檢查歸檔資料夾是否已經有同名檔案，避免 File.Move 發生衝突報錯
                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath);
                }

                // 將檔案搬移到歸檔資料夾
                File.Move(file, destinationPath);
            }
        }

        private StringBuilder Makeup_IntH1B104_SQL(string strLine, MIFUtil _mesdb, ILog _logGET)
        {
            // extract the content
            string strProbeCardId = string.Empty;
            string strAccName = _mesdb.SubStr(strLine, 0, 30).Trim();
            string strAccId = _mesdb.SubStr(strLine, 30, 10).Trim();
            strProbeCardId = strAccName + "-" + strAccId;
            string strCreateTime = util.SubStr(strLine, 40, 18).Trim();
            string strCODE = util.SubStr(strLine, 58, 20).Trim();
            string strRepairReason = util.SubStr(strLine, 78, 50).Trim();
            string strRepairProcess = util.SubStr(strLine, 128, 50).Trim();
            string strPcdmsReason = util.SubStr(strLine, 178, 25).Trim();
            string strPcdmsUserID = _mesdb.SubStr(strLine, 203, 5).Trim();
            string strTestMode = _mesdb.SubStr(strLine, 263, 10).Trim();
            string strTemperature = _mesdb.SubStr(strLine, 273, 10).Trim();
            string strEquipment = _mesdb.SubStr(strLine, 283, 10).Trim();
            string strOverDriver = _mesdb.SubStr(strLine, 293, 10).Trim();
            string strAbnormalDesc = _mesdb.SubStr(strLine, 333, 300).Trim();

            //TBL_ACC_REPAIR_TEMP table trigger : trg_acc_repair_temp
            /*若傳回的ProbeCardID code='REPAIR_OK', 而且其狀態為REPAIR，才自動更新為IDLE;
              其餘狀態皆不允許變動。
              (並新增一筆Tbl_Acc_History記錄; TBL_ACC_REPAIR_TEMP.DeleteFlag := 'Y').
            */

            StringBuilder sSql = new StringBuilder();

            sSql.AppendFormat(@"UPDATE fwadmin.TBL_ACC_REPAIR_TEMP SET CODE='{0}', 
                                      REPAIRREASON='{1}', 	--M200705044新增欄位
                                      REPAIRPROCESS='{2}',  --M200705044新增欄位
                                      PCDMS_Reason='{3}', 
                                      PCDMS_UserID='{4}', 
                                      TESTMODE = '{5}', --PCD 健康管理新增欄位
                                      TEMPERATURE = '{6}', --PCD 健康管理新增欄位
                                      TESTERID = '{7}', --PCD 健康管理新增欄位
                                      OVERDRIVER = '{8}', --PCD 健康管理新增欄位
                                      ABNORMAL_DESC = '{9}' --PCD 健康管理新增欄位
                                      WHERE PROBECARDID='{10}' AND 
                                            CREATETIME='{11}'"
                , strCODE, strRepairReason, strRepairProcess, strPcdmsReason, strPcdmsUserID
                , strTestMode, strTemperature, strEquipment, strOverDriver, strAbnormalDesc
                , strProbeCardId, strCreateTime);

            _logGET.Info("strSQL:" + sSql.ToString());
            return sSql;
        }
    }
}
