	public HashMap[] getSubjectData(Connection con,String year , String month,
												   String dept_id,String station_id,
												   String title,String emp_table,
												   String[] shift_id ,
												   String loginUserID,String funcName) throws Exception {
		HashMap[] hm = null;	
		String shift;
		String sql =" select distinct a.emp_id, a.item, a.detailitem, a.record, a.comments  ";
		  sql =sql +" from rbl_dl_performance_subject a , "+emp_table+" b ";
		  sql =sql +" ,rbl_dl_item c ,rbl_dl_detailitem d   ";
		  sql =sql +" where a.year = '"+year+"' and a.month = '"+month+"' ";
		  sql =sql +" and a.emp_id = b.emp_id and b.title ='"+title+"' and b.station_id ='"+station_id+"' and b.dept_id='"+dept_id+"' ";
		  if(shift_id.length >0){
				shift="";
				for(int i=0;i<shift_id.length;i++){
					if(shift.length()==0){
						shift = shift +"'"+shift_id[i]+"'";
					}
					else{
						shift = shift +",'"+shift_id[i]+"'";
					}					
				}
				if(shift.length()>0){				
					sql=sql+ " and b.shift_id in ( "+ shift + ")";
				}
			}		  
		  sql =sql +" and b.title = c.title(+) and b.station_id = c.station_id(+) ";
		  sql =sql +" and c.station_id=d.station_id(+) and c.title=d.title(+) and c.item=d.item(+) and c.subjective(+)='Y'  ";
		  sql =sql +" and c.deleteflag(+)='N' and d.deleteflag(+)='N' ";
		  sql =sql +" order by a.emp_id,a.item,a.detailitem ";
		  //System.out.println(sql);
		hm = SESDB.qryHashMapBySql(con,sql,new Object[]{});	
		try{
			DBAcc db = new DBAcc(con);
			int icheck = db.insertconfidential("DL", loginUserID, funcName, sql);
		}
		catch (Exception e){
			e.printStackTrace();
		}	
		return hm;		
	}	

public IEnumerable<dynamic> GetSubjectData(
        IDbConnection con, string year, string month, 
        string deptId, string stationId, string title, 
        string empTable, string[] shiftIds, 
        string loginUserId, string funcName)
    {
        IEnumerable<dynamic> result = null;

        // 使用 $@ 宣告多行字串，empTable 無法參數化，故直接內插拼接
        string sql = $@"
            SELECT DISTINCT a.emp_id, a.item, a.detailitem, a.record, a.comments  
            FROM rbl_dl_performance_subject a, {empTable} b, rbl_dl_item c, rbl_dl_detailitem d   
            WHERE a.year = :Year 
              AND a.month = :Month 
              AND a.emp_id = b.emp_id 
              AND b.title = :Title 
              AND b.station_id = :StationId 
              AND b.dept_id = :DeptId 
              AND b.title = c.title(+) 
              AND b.station_id = c.station_id(+) 
              AND c.station_id = d.station_id(+) 
              AND c.title = d.title(+) 
              AND c.item = d.item(+) 
              AND c.subjective(+) = 'Y'  
              AND c.deleteflag(+) = 'N' 
              AND d.deleteflag(+) = 'N' ";

        // 使用 DynamicParameters 加入參數 (防止 SQL Injection)
        var parameters = new DynamicParameters();
        parameters.Add("Year", year);
        parameters.Add("Month", month);
        parameters.Add("Title", title);
        parameters.Add("StationId", stationId);
        parameters.Add("DeptId", deptId);

        // 如果 shiftIds 陣列有值，直接加入 Dapper 參數，Dapper 會自動處理 IN 裡面的展開
        if (shiftIds != null && shiftIds.Length > 0)
        {
            sql += " AND b.shift_id IN :ShiftIds ";
            parameters.Add("ShiftIds", shiftIds);
        }

        sql += " ORDER BY a.emp_id, a.item, a.detailitem";

        try
        {
            // Dapper 執行查詢，回傳的動態物件集合相當於原本的 HashMap[]
            result = con.Query(sql, parameters);

            // 模擬原本 Java 裡的 log 紀錄，請替換為 C# 專案中負責寫 Log 的服務方法
            // DBAcc db = new DBAcc(con);
            // db.InsertConfidential("DL", loginUserId, funcName, sql);
        }
        catch (Exception ex)
        {
            // 記錄或拋出例外
            Console.WriteLine(ex.ToString());
            throw;
        }

        return result;
    }
