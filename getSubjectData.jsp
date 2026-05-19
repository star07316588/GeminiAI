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

	public HashMap[] getTotalData_ph(Connection con,String startyear,String startmonth,String endtyear,String endtmonth,
			  String station,String shift,String deptId,String empId,String DateType,String section,String userid,String func) throws Exception{
		ArrayList al = new ArrayList();
		al.add(startyear+startmonth);
		al.add(endtyear+endtmonth);
		String sql_ins="";
		String sql = 			
		"select ps.year,\n" +
		"       ps.month,\n" + 
		"       se.dept_id,\n" + 
		"       se.station_id,\n" + 
		"       se.shift_id,\n" + 
		"       se.emp_id,\n" + 
		"       se.name,\n" + 
		"       se.position_group,\n" + 
		"       se.title,\n" + 
		"       to_char(se.arrive_date,'yyyy/mm/dd') arrive_date,\n" + 
		"       ps.score RANKING,\n" + 
		"       ps.ranking_shift,\n" + 
		"       ps.ranking_shiftA,\n" + 
		"       ps.Ranking_Title,\n" + 
		"       ps.Ranking_Group1,\n";
				
		sql =sql + "ps.bonus,ps.balance ";
		
		sql =sql +	
		"  from Rbl_DL_Performance_Summary ps,\n" + 
		"       Sbl_Emp                    se,\n" + 			
		"		rbl_dl_organization org " +
		" where ps.emp_id = se.emp_id  " ;
		sql_ins =sql;
		sql+=" and ps.year || ps.month between ? and ?\n";
		sql_ins+=" and ps.year || ps.month between '"+startyear+startmonth+"' and '"+endtyear+endtmonth+"'\n";
		
		if(station!=null && !station.equals("")){
		sql +="	and se.station_id = ?\n";
		al.add(station);
		sql_ins+="	and se.station_id = '"+station+"'\n";
		}
		if(shift!=null && !shift.equals("")){
		sql +="	and se.shift_id = ?\n";
		al.add(shift);
		sql_ins+="	and se.shift_id = '"+shift+"'\n";
		}
		if(deptId!=null && !deptId.equals("")){
		sql +="	and se.dept_id = ?\n";
		al.add(deptId);
		sql_ins+="	and se.dept_id = '"+deptId+"'\n";
		}
		if(empId!=null && !empId.equals("")){
		sql +="	and se.emp_id = ?\n";
		al.add(empId);
		sql_ins+="	and se.emp_id = '"+empId+"'\n";
		}		
		if(DateType!=null && !DateType.equals("")){
		sql +="	and instr(ps.month,'"+DateType+"') > 0  \n";	
		sql_ins+="	and instr(ps.month,'"+DateType+"') > 0  \n";
		}		
		sql += " and se.station_id = org.station_id \n"+ 
		" and se.shift_id = org.shift_id \n"+ 
		" and se.dept_id = org.dept_id \n"+
		" and org.section ='"+section+"' \n" +		
		" order by ps.year,\n" + 
		"          ps.month,\n" + 
		"          se.shift_id,\n" + 
		"          se.station_id,\n" + 
		"          se.title,\n" + 
		"          se.emp_id";
		sql_ins += " and se.station_id = org.station_id \n"+ 
				" and se.shift_id = org.shift_id \n"+ 
				" and se.dept_id = org.dept_id \n"+
				" and org.section ='"+section+"' \n" +		
				" order by ps.year,\n" + 
				"          ps.month,\n" + 
				"          se.shift_id,\n" + 
				"          se.station_id,\n" + 
				"          se.title,\n" + 
				"          se.emp_id";
		//System.out.print(sql);	
		try{
			DBAcc da =new DBAcc(con);   
			da.insertconfidential("DL", userid, func, sql_ins);
		}
		catch(Exception e){
			e.printStackTrace();	
		}
		return  SESDB.qryHashMapBySql(con,sql,(Object[])al.toArray(new Object[0]));
	}
