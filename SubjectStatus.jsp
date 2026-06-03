		sSQL="select a.status from Rbl_DL_Status a where year='"+year+"' and month='"+month+"' and dept_id='"+dept_id+"' and title='"+title+"' and type='SUBJECT' ";
		if(station_id.length()>0){	
			s_id=" ( ";
			String[] StrArray=station_id.split(",");
			  for(int i=0;i<StrArray.length;i++)
			  {				 
				  s_id=s_id + " '"+StrArray[i]+"' ,";
			  }
			  
			  s_id = s_id.substring(0,s_id.length()-1)+" ) ";			
			sSQL+= "and station_id in "+s_id+" ";
		}	
		if(shift_id.length()>0){	
			Shiftids=" ( ";
			String[] StrArray=shift_id.split(",");
			  for(int i=0;i<StrArray.length;i++)
			  {				 
				  Shiftids=Shiftids + " '"+StrArray[i]+"' ,";
			  }
			  
			  Shiftids = Shiftids.substring(0,Shiftids.length()-1)+" ) ";			
			sSQL+= "and shift_id in "+Shiftids+" ";
		}
		//System.out.println(sSQL);
		rs = jdbc.queryData(sSQL,con);	

// 因為 SQL 只 Select 了單一欄位 (status)，我們可以直接回傳 IEnumerable<string>
    public IEnumerable<string> GetSubjectStatus(
        IDbConnection con, string year, string month, 
        string deptId, string title, string stationId, string shiftId)
    {
        // 宣告 DynamicParameters 來集中管理參數
        var parameters = new DynamicParameters();
        parameters.Add("Year", year);
        parameters.Add("Month", month);
        parameters.Add("DeptId", deptId);
        parameters.Add("Title", title);

        // 使用 StringBuilder 建立基礎 SQL (使用逐字字串 $@ 保持排版乾淨)
        var sqlBuilder = new StringBuilder($@"
            SELECT a.status 
            FROM Rbl_DL_Status a 
            WHERE year = :Year 
              AND month = :Month 
              AND dept_id = :DeptId 
              AND title = :Title 
              AND type = 'SUBJECT' 
        ");

        // 處理 station_id：直接切成陣列，丟給 Dapper 處理 IN 語句
        if (!string.IsNullOrEmpty(stationId))
        {
            string[] stationArray = stationId.Split(',');
            sqlBuilder.AppendLine(" AND station_id IN :StationIds ");
            parameters.Add("StationIds", stationArray);
        }

        // 處理 shift_id：同樣切成陣列，丟給 Dapper
        if (!string.IsNullOrEmpty(shiftId))
        {
            string[] shiftArray = shiftId.Split(',');
            sqlBuilder.AppendLine(" AND shift_id IN :ShiftIds ");
            parameters.Add("ShiftIds", shiftArray);
        }

        try
        {
            // Dapper 執行查詢，因為只抓 status 欄位，直接 mapping 到字串集合
            return con.Query<string>(sqlBuilder.ToString(), parameters);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            throw;
        }
    }

            var empStatus = _reportService.GetSubjectStatus(searchInfo.Year, searchInfo.Month, searchInfo.Title, searchInfo.Dept_Id, Station_Id, searchInfo.Shift_Id);
            foreach (string status in empStatus)
            {
                if(searchInfo.LoginTitle == "LEADER")
                {
                    if(status != "LEADER PROCESSING")
                    {
                        bcheck = false;
                    }
                }
                else if (searchInfo.LoginTitle == "SUPERVISOR")
                {
                    if(status != "LEADER OK")
                    {
                        bcheck = false;
                    }
                }
                else if (searchInfo.LoginTitle == "MANAGER")
                {
                    if(status != "SUPERVISOR OK")
                    {
                        bcheck = false;
                    }
                }
                else
                {
                    bcheck = false;
                }
            }
            return Content(JsonConvert.SerializeObject(result), "application/json");
