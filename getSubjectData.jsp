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
