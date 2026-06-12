package com.mxic.dl.tool;

import java.sql.*;
import java.util.*;
import com.mxic.dl.tool.*;

public class Month_Cal_2 {	
	public Month_Cal_2(){
		
	}
//總分/總獎金/餘額整理 Start	
	public int getManagerDatatoSummary(String dept_id,String section ,Connection conn)throws Exception{
		int result=0;
		Statement stmt = null;
		Statement stmt2 = null;
		ResultSet rs;
		ResultSet rs2;
		String sSQL="";		
		String emp_id="";
		String station_id="";
		String year="";
		String month="";
		String score="";
		String bonus="";
		String balance="";
		String title="";
		int maximum=0;
		int insertcheck=0;	
		java.util.LinkedList sql_insert = new java.util.LinkedList();			
		try{
				stmt = conn.createStatement();		
				stmt2= conn.createStatement();
				/*sSQL="select distinct a.emp_id, a.year, a.month,b.title " +
					 " from rbl_dl_performance_manager a, sbl_emp b "+
					 " where a.emp_id = b.emp_id and b.dept_id ='"+dept_id+"' "+
					 " and a.YEAR || a.MONTH =to_char(add_months(sysdate,-1 ), 'YYYYMM') ";	*/
				//Modify by Sam on 20130308,增加station_id欄位,section條件
				sSQL="select distinct a.emp_id, a.year, a.month,b.title ,b.station_id " +
				 " from rbl_dl_performance_manager a, " +
				 " (select distinct b1.emp_id,b1.title,b1.dept_id,b1.station_id from rbl_dl_emp b1,rbl_dl_organization b2    "+
				 " 		where  b1.dept_id = b2.dept_id and b1.shift_id=b2.shift_id and b1.station_id = b2.station_id "+
				 "		and b1.dept_id ='"+dept_id+"' and b2.section ='"+section+"'  "+
				 " 		)  b "+
				 " where a.emp_id = b.emp_id  "+
				 " and a.YEAR || a.MONTH =to_char(add_months(sysdate,-1 ), 'YYYYMM') ";	
				stmt.executeQuery(sSQL);				
				rs = stmt.getResultSet();								
	            while(rs.next()){
	            	
	            	emp_id = rs.getString("emp_id");
	            	year = rs.getString("year");
	            	month = rs.getString("month");
	            	//Add by Sam on 20120319 ,增加取設定的上限金額
	            	title = rs.getString("title");	     
	            	station_id =rs.getString("station_id");
	            	//Modify by Sam on 20130308,增加station_id條件
	            	sSQL="select nvl(a.maximum,0) as maximum " +
            		" from rbl_dl_bonusbase a "+
            		" where a.title ='"+title+"' and a.station_id ='"+station_id+"'  and a.deleteflag ='N' ";		            	
					stmt2.executeQuery(sSQL);				
					rs2 = stmt2.getResultSet();					
					if (rs2.next()){
						maximum=rs2.getInt("maximum");
					}else
						maximum =0;
	            	
	            	sSQL="select nvl(a.final,0) as final " +
	            		" from rbl_dl_performance_manager a "+
	            		" where a.emp_id ='"+emp_id+"' and a.year || a.month ='"+year+month+"' "+ 
	            		" and a.item like '經理評比%Ranking' ";		            	
					stmt2.executeQuery(sSQL);				
					rs2 = stmt2.getResultSet();					
					if (rs2.next()){
						score=rs2.getString("final");
					}else
						score ="0";
					//Modify by Sam on 20140416 for ReqNo:JC201400095,增加balance欄位
					sSQL="select nvl(a.total,0) as total ,nvl(a.record,0) as record ,nvl(a.final,0) as final , nvl(a.balance,0) as balance " +
            		" from rbl_dl_performance_manager a "+
            		" where a.emp_id ='"+emp_id+"' and a.year || a.month ='"+year+month+"' "+ 
            		" and a.item like '經理評比%績效' ";		            	
					stmt2.executeQuery(sSQL);				
					rs2 = stmt2.getResultSet();					
					if (rs2.next()){
						//Add by Sam on 20140416 for ReqNo:JC201400095 改為Bonus = final ,balance = balance 
						bonus=rs2.getString("final");
						balance=rs2.getString("balance");
						/* Mark by Sam on 20140416 for ReqNo:JC201400095 ,取消這些判斷
						if (rs2.getDouble("total")>=0){
							//Modify by Sam on 20120319,Bonus>6400時以6400算,餘額列入Balance
							
							if(Integer.parseInt(rs2.getString("final"))>maximum){
								bonus=Integer.toString(maximum);
								balance =Integer.toString(Integer.parseInt(rs2.getString("final")) - maximum);
							}
							else{
								bonus=rs2.getString("final");
								balance="0";
							}
						}
						else{
							bonus=rs2.getString("record");
							balance=rs2.getString("total");
						} */
					}else{
						bonus="0";
						balance="0";
					}														
            		sSQL="insert into Rbl_DL_Performance_summary(emp_id,year,month,score,bonus,balance,createuserid,createtime) values (?,?,?,?,?,?,'SYS',to_char(sysdate,'yyyymmdd hh24miss')||'000')";
            		sql_insert.clear();
            		sql_insert.add(sSQL);
            		sql_insert.add(emp_id);
            		sql_insert.add(year);
            		sql_insert.add(month);
            		sql_insert.add(score);
            		sql_insert.add(bonus);
            		sql_insert.add(balance);            			            	
            		insertcheck = insert(sql_insert,conn);	            		
	            	            			            
	            }	         	              
	            result = 1;
		    }catch(Exception e){		    	
		    	e.printStackTrace();
		    	result = 0;
		    }finally{
		    	if(stmt != null)
					try {
						stmt.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}
		    	if(stmt2 != null)
					try {
						stmt2.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}		
		    }
		    return result;		
	}
	//總分/總獎金/餘額整理 End
		
	//總分排名計算 start
	//Modify by Sam on 20111221,增加Shift_id排名A,B,C,D...etc
	public int getSummaryRanking(String dept_id,String section,Connection conn)throws Exception{
		int result=0;
		Statement stmt = null;
		ResultSet rs;
		String sSQL="";	
		String emp_id="";
		String year="";
		String month="";
		String ranking_shift="";
		String ranking_station="";
		String ranking_title="";
		String ranking_group1="";
		String ranking_groupa="";
		String ranking_shifta="";
		java.util.LinkedList sql_update = new java.util.LinkedList();
		int updatecheck=0;
		try{
			stmt = conn.createStatement();			
			/*sSQL = " select a.emp_id,a.year,a.month  "+
					"        ,rank() over(partition by b.shift_id, b.title order by nvl(a.score, 0) desc) as ranking_shift "+
					"        ,rank() over(partition by b.shift_id, b.title,b.station_id order by nvl(a.score, 0) desc) as ranking_station       "+
					"        ,rank() over(partition by b.title order by nvl(a.score, 0) desc) as ranking_title "+
					"        ,rank() over(partition by b.title,b.position_group order by nvl(a.score, 0) desc) as ranking_group1 "+
					"        ,fun_get_ranking_level(a.emp_id,b.title,b.position_group,b.dept_id,a.year,a.month) as ranking_groupa "+
					"        ,fun_get_shift_ranking_level(a.emp_id,b.title,b.shift_id,b.dept_id,a.year,a.month) as ranking_shifta "+
					"   from dlowner.rbl_dl_performance_summary a, sbl_emp b "+
					"  where a.emp_id = b.emp_id and a.score is not null and b.dept_id='"+dept_id+"' "+
					" and a.YEAR || a.MONTH =to_char(add_months(sysdate,-1 ), 'YYYYMM') ";*/
			//Modify by Sam on 20130308,增加section條件
			/*sSQL = " select a.emp_id,a.year,a.month  "+
				"        ,rank() over(partition by b.shift_id, b.title order by nvl(a.score, 0) desc) as ranking_shift "+
				"        ,rank() over(partition by b.shift_id, b.title,b.station_id order by nvl(a.score, 0) desc) as ranking_station       "+
				"        ,rank() over(partition by b.title order by nvl(a.score, 0) desc) as ranking_title "+
				"        ,rank() over(partition by b.title,b.position_group order by nvl(a.score, 0) desc) as ranking_group1 "+
				"        ,fun_get_ranking_level(a.emp_id,b.title,b.position_group,b.dept_id,a.year,a.month) as ranking_groupa "+
				"        ,fun_get_shift_ranking_level(a.emp_id,b.title,b.shift_id,b.dept_id,a.year,a.month) as ranking_shifta "+
				"   from dlowner.rbl_dl_performance_summary a, (select distinct emp_id,title,station_id,shift_id,dept_id,position_group from rbl_dl_performance_item where year || month  =to_char(add_months(sysdate,-1 ), 'YYYYMM')   )  b "+
				"  where a.emp_id = b.emp_id and a.score is not null and b.dept_id='"+dept_id+"' " + 
				" and a.YEAR || a.MONTH =to_char(add_months(sysdate,-1 ), 'YYYYMM') "; */
			sSQL = " select a.emp_id,a.year,a.month  "+
					"        ,rank() over(partition by b.shift_id, b.title order by nvl(a.score, 0) desc) as ranking_shift "+
					"        ,rank() over(partition by b.shift_id, b.title,b.station_id order by nvl(a.score, 0) desc) as ranking_station       "+
					"        ,rank() over(partition by b.title order by nvl(a.score, 0) desc) as ranking_title "+
					"        ,rank() over(partition by b.title,b.position_group order by nvl(a.score, 0) desc) as ranking_group1 "+
					"        ,fun_get_ranking_level(a.emp_id,b.title,b.position_group,b.dept_id,a.year,a.month,'"+section+"') as ranking_groupa "+
					"        ,fun_get_shift_ranking_level(a.emp_id,b.title,b.shift_id,b.dept_id,a.year,a.month,'"+section+"') as ranking_shifta "+
					"   from dlowner.rbl_dl_performance_summary a, "+ 
					"   (select distinct b1.emp_id,b1.title,b1.station_id,b1.shift_id,b1.dept_id,b1.position_group "+
					"    from rbl_dl_emp b1 , rbl_dl_organization b2 "+
					"    where b1.dept_id =b2.dept_id and b1.shift_id=b2.shift_id and b1.station_id=b2.station_id "+
					"	 and b1.dept_id='"+dept_id+"' and b2.section ='"+section+"' "+
					"    )  b "+
					"  where a.emp_id = b.emp_id and a.score is not null  " + 
					" and a.YEAR || a.MONTH =to_char(add_months(sysdate,-1 ), 'YYYYMM') ";			
			
			stmt.executeQuery(sSQL);				
			rs = stmt.getResultSet();								
	        while(rs.next()){	  
            	ranking_shift=rs.getString("ranking_shift");	            		
            	ranking_station=rs.getString("ranking_station");
            	ranking_title=rs.getString("ranking_title");
            	ranking_group1=rs.getString("ranking_group1");
            	ranking_groupa=rs.getString("ranking_groupa");
            	ranking_shifta=rs.getString("ranking_shifta");
            	emp_id=rs.getString("emp_id");
            	year=rs.getString("year");
            	month=rs.getString("month");   
            	
	        	sSQL="update Rbl_DL_Performance_summary set ranking_shift=?,ranking_station=?,ranking_title=?,ranking_group1=?,ranking_groupa=? ,ranking_shifta=?　 " +
	        		" where emp_id=? and year=? and month=? ";
        		sql_update.clear();	            	
        		sql_update.add(sSQL);
        		sql_update.add(ranking_shift);
        		sql_update.add(ranking_station);
        		sql_update.add(ranking_title);
        		sql_update.add(ranking_group1);
        		sql_update.add(ranking_groupa);
        		sql_update.add(ranking_shifta);
        		sql_update.add(emp_id);
        		sql_update.add(year);
        		sql_update.add(month);     			            		
        		updatecheck = update(sql_update,conn);	        		
	        }
	        result = 1;
		}catch(Exception e){	
			e.printStackTrace();	      
			result = 0;
		}finally{
			if(stmt != null)
				try {
					stmt.close();
				} catch (SQLException e) {					
					e.printStackTrace();
				}	    	
		}	
		return result;
	}
	//總分排名計算 End
	//Add by Sam on 20121203,JC201200334,by 季 Update Rbl_DL_Performance_summary.SkipFlag
	public int SetSummarySkipFlag(String dept_id,String section,Connection conn)throws Exception{
		int result=0;
		Statement stmt = null;	
		ResultSet rs;		
		String sSQL="";		
		String emp_id="";
		String year="";
		String month="";
		int updatecheck=0;
		java.util.LinkedList sql_update = new java.util.LinkedList();
		/*sSQL=" select distinct a.emp_id ,a.year,a.month from dlowner.rbl_dl_performance_summary a, sbl_emp b "+
		 "  where a.emp_id = b.emp_id and b.dept_id='"+dept_id+"' "+
		 "            and a.year = to_char(add_months(sysdate, -1), 'yyyy')"+
		 "            and a.year || a.month >= to_char(add_months(sysdate, -3), 'yyyymm')"+
		 "            and a.year || a.month <= to_char(add_months(sysdate, -1), 'yyyymm')"+			 
		 " ";*/
		//Modify by Sam on 20130311,增加rbl_dl_organization section條件 
		sSQL=" select distinct a.emp_id ,a.year,a.month from dlowner.rbl_dl_performance_summary a, "+
		 " (select distinct b1.emp_id,b1.title,b1.station_id,b1.shift_id,b1.dept_id,b1.position_group "+
		 "  from rbl_dl_emp b1, rbl_dl_organization b2 " +
		 "  where b1.dept_id = b2.dept_id and b1.shift_id=b2.shift_id and b1.station_id = b2.station_id "+
		 "  and  b1.dept_id ='"+dept_id+"' AND b2.Section ='"+section+"' "+
		 "  )  b "+
		 "  where a.emp_id = b.emp_id and b.dept_id='"+dept_id+"' "+
		 "            and a.year = to_char(add_months(sysdate, -1), 'yyyy')"+
		 "            and a.year || a.month >= to_char(add_months(sysdate, -3), 'yyyymm')"+
		 "            and a.year || a.month <= to_char(add_months(sysdate, -1), 'yyyymm')"+			 
		 " ";		
		try{
			stmt = conn.createStatement();
			stmt.executeQuery(sSQL);				
			rs = stmt.getResultSet();		
			while(rs.next()){	   
				emp_id = rs.getString("emp_id");
            	year = rs.getString("year");
            	month = rs.getString("month");            					
				sSQL="update Rbl_DL_Performance_summary T " +
					"  SET SKIPFLAG = DECODE((SELECT DECODE(?,'01',JAN,'02',FEB,'03',MAR,'04',APR,'05',MAY,'06',JUN," + 
					" '07',JUL,'08',AUG,'09',SEP,'10',OCT,'11',NOV,'12',DEC," +
					" 'N') FROM RBL_DL_RANKINGSET WHERE EMP_ID = T.EMP_ID AND YEAR = T.YEAR) ,'X','Y','N') "+
					" where emp_id=? and year=? and month=? ";
				sql_update.clear();
				sql_update.add(sSQL);
				sql_update.add(month);
				sql_update.add(emp_id);					
				sql_update.add(year);
				sql_update.add(month);
				updatecheck = update(sql_update,conn);	        							
			}
			result = 1;
		}catch(Exception e){	
			e.printStackTrace();
	    	result = 0;
		}finally{
			if(stmt != null)
				try {
					stmt.close();
				} catch (SQLException e) {					
					e.printStackTrace();
				}	
		}						
		return result;
	}
	
	//半年總分排名計算 Start
	//Modify by Sam on 20111221,增加Shift_id排名A,B,C,D...etc
	public int sixMonthRanking(String dept_id,String section , Connection conn)throws Exception{
		int result=0;
		Statement stmt = null;	
		Statement stmt2 = null;
		Statement stmt_emp = null;
		ResultSet rs;		
		ResultSet rs2;
		ResultSet rs_emp;
		String sSQL="";		
		String emp_id="";
		String year="";
		String month="";
		String score="";
		String bonus="";
		String balance="";
		String ranking_shift="";
		String ranking_station="";
		String ranking_title="";
		String ranking_group1="";
		String ranking_groupa="";
		String ranking_shifta="";
		int insertcheck=0;	
		java.util.LinkedList sql_insert = new java.util.LinkedList();			
		try{
				stmt = conn.createStatement();	
				stmt2 = conn.createStatement();	
				 //Add by Sam  on 20130715 for ReqNo:JC201300204 ,人員名單從原本取得分數的來源抽出單獨取得,將人員資料傳入分開計算分數與獎金
				stmt_emp = conn.createStatement();
				sSQL = "   select distinct a1.emp_id,a1.year, case when to_char(add_months(sysdate, -1), 'mm') <= '06' then '1H' else '2H' end as month " +
	                   "   from rbl_dl_performance_summary a1 , " +
		               "   (select distinct bb1.emp_id,bb1.title,bb1.station_id,bb1.shift_id,bb1.dept_id,bb1.position_group   " +
		               "   from rbl_dl_emp bb1,rbl_dl_organization bb2  " +
		               "   where bb1.dept_id = bb2.dept_id and bb1.shift_id = bb2.shift_id and bb1.station_id = bb2.station_id " +
		               "   and bb1.dept_id ='"+dept_id+"' AND bb2.section ='"+section+"' " +
		               "   ) b1 " +
	                   "   where a1.emp_id = b1.emp_id and b1.dept_id ='"+dept_id+"'      " +
	                   "   and a1.year = to_char(add_months(sysdate, -1), 'yyyy')" +
	                   "   and a1.year || a1.month >= to_char(add_months(sysdate, -6), 'yyyymm') " +
	                   "   and a1.year || a1.month <= to_char(add_months(sysdate, -1), 'yyyymm') ";
				stmt_emp.executeQuery(sSQL);				
				rs_emp = stmt_emp.getResultSet();	
				while(rs_emp.next()){
					
					emp_id = rs_emp.getString("emp_id");
	            	year = rs_emp.getString("year");
	            	month = rs_emp.getString("month");
	            	//System.out.println("EMP:"+emp_id);
					//Modify by Sam on 20120312,Bonus從avg改為Sum,Balance改為取最後一個月的值
					//Modify by Sam on 20121203,JC201200344 ,增加判斷 rbl_dl_performance_summary.SkipFlag = 'N' ,Bonus及balance獨立運算							
					sSQL=" select emp_id , year,month ,title , score , ranking_shift,ranking_station,ranking_title,ranking_group1,ranking_groupa,ranking_shifta from " + 
					    " ( select a.emp_id,a.year,b.title,"+
						"        case when to_char(add_months(sysdate, -1), 'mm') <= '06' then '1H' else '2H' end as month ,a.score ,"+
						"        decode(a.score,null,null,rank() over(partition by b.shift_id, b.title order by nvl(a.score, 0) desc)) as ranking_shift,"+
						"        decode(a.score,null,null,rank() over(partition by b.shift_id, b.title, b.station_id order by nvl(a.score, 0) desc)) as ranking_station,"+
						"        decode(a.score,null,null,rank() over(partition by b.title order by nvl(a.score, 0) desc)) as ranking_title,"+
						"        decode(a.score,null,null,rank() over(partition by b.title, b.position_group order by nvl(a.score, 0) desc)) as ranking_group1,"+
						"        decode(a.score,null,null,fun_get_ranking_level(a.emp_id, b.title, b.position_group,b.dept_id,a.year,case when to_char(add_months(sysdate, -1), 'mm') <= '06' then '1H' else '2H' end,'"+section+"' )) as ranking_groupa,"+
						"        decode(a.score,null,null,fun_get_shift_ranking_level(a.emp_id, b.title, b.shift_id,b.dept_id,a.year,case when to_char(add_months(sysdate, -1), 'mm') <= '06' then '1H' else '2H' end,'"+section+"' )) as ranking_shifta"+
						"   from (select a1.emp_id,a1.year,round(avg(score),2) as score,round(sum(bonus)) as bonus,round(avg(balance),2) as balance "+
						"           from rbl_dl_performance_summary a1 ,"+
						"			(select distinct bb1.emp_id,bb1.title,bb1.station_id,bb1.shift_id,bb1.dept_id,bb1.position_group  "+ 
						"			from rbl_dl_emp bb1,rbl_dl_organization bb2  "+
						"			where bb1.dept_id = bb2.dept_id and bb1.shift_id = bb2.shift_id and bb1.station_id = bb2.station_id " +
						"			and bb1.dept_id ='"+dept_id+"' AND bb2.section ='"+section+"' "+
						" 			) b1 "+
						"          where a1.emp_id = b1.emp_id and b1.dept_id ='"+dept_id+"'      "+
						"            and a1.year = to_char(add_months(sysdate, -1), 'yyyy')"+
						"            and a1.year || a1.month >= to_char(add_months(sysdate, -6), 'yyyymm')"+
						"            and a1.year || a1.month <= to_char(add_months(sysdate, -1), 'yyyymm')"+		
						"            and a1.skipflag = 'N' "+		
						"          group by a1.emp_id,a1.year ) a,   (select distinct emp_id,title,station_id,shift_id,dept_id,position_group from rbl_dl_emp ) b"+			
						"          where a.emp_id = b.emp_id ) " + 
						" where emp_id='"+emp_id+"'  ";					
					stmt.executeQuery(sSQL);				
					rs = stmt.getResultSet();								
		            if(rs.next()){	            	
		            	//emp_id = rs.getString("emp_id");
		            	//year = rs.getString("year");
		            	//month = rs.getString("month");
		            	score = rs.getString("score");
		            	//bonus = rs.getString("bonus");
		            	//balance = rs.getString("balance");
		            	ranking_shift=rs.getString("ranking_shift");	            		
		            	ranking_station=rs.getString("ranking_station");
		            	ranking_title=rs.getString("ranking_title");
		            	ranking_group1=rs.getString("ranking_group1");
		            	ranking_groupa=rs.getString("ranking_groupa");			
		            	ranking_shifta=rs.getString("ranking_shifta");		            
		            }else{
		            	score ="";
		            	ranking_shift ="";
		            	ranking_station ="";
		            	ranking_title ="";
		            	ranking_group1 ="";
		            	ranking_groupa ="";
		            	ranking_shifta="";
		            }	
		            //System.out.println("EMP:"+emp_id + " score:" + score  );
	            	//Modify by Sam on 20121205,JC201200344 BONUS跟BALANCE獨立取得,不受SkipFlag限制
	            	sSQL=" select a.emp_id ,a.bonus,nvl(c.balance,0) as balance " +
            		" from (select a1.emp_id, round(sum(nvl(bonus,0))) as bonus  from rbl_dl_performance_summary a1 "+
            		" 			where a1.emp_id ='"+emp_id+"' "+ 
					"            and a1.year = to_char(add_months(sysdate, -1), 'yyyy')"+
					"            and a1.year || a1.month >= to_char(add_months(sysdate, -6), 'yyyymm')"+
					"            and a1.year || a1.month <= to_char(add_months(sysdate, -1), 'yyyymm') group by a1.emp_id ) a "+
					" ,(Select c1.emp_id,c1.balance from rbl_dl_performance_summary c1 where c1.year || c1.month = to_char(add_months(sysdate, -1), 'yyyymm')  ) c "+
					" where a.emp_id = c.emp_id(+) ";	            	
					stmt2.executeQuery(sSQL);				
					rs2 = stmt2.getResultSet();					
					if (rs2.next()){
						bonus = rs2.getString("bonus");
						balance = rs2.getString("balance");
					}else{
						bonus = "0";
						balance ="0";
					}				
					//System.out.println("EMP:"+emp_id + " bonus:" + bonus + " balance:" + balance );
            		sSQL="insert into Rbl_DL_Performance_summary(emp_id,year,month,score,bonus,balance,"+
            			 " ranking_shift,ranking_station,ranking_title,ranking_group1,ranking_groupa,ranking_shifta,"+
            			 " createuserid,createtime) values (?,?,?,?,?,?,?,?,?,?,?,?,"+
            			 "'SYS',to_char(sysdate,'yyyymmdd hh24miss')||'000')";
            		sql_insert.clear();
            		sql_insert.add(sSQL);
            		sql_insert.add(emp_id);
            		sql_insert.add(year);
            		sql_insert.add(month);
            		sql_insert.add(score);
            		sql_insert.add(bonus);
            		sql_insert.add(balance);     
            		sql_insert.add(ranking_shift);
            		sql_insert.add(ranking_station);
            		sql_insert.add(ranking_title);
            		sql_insert.add(ranking_group1);
            		sql_insert.add(ranking_groupa);
            		sql_insert.add(ranking_shifta);
            		insertcheck = insert(sql_insert,conn);	            				            	            			            		            	     					
				}
				    	              
	            result = 1;
		    }catch(Exception e){		    	
		    	e.printStackTrace();
		    	result = 0;
		    }finally{
		    	if(stmt != null)
					try {
						stmt.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}
			    if(stmt2 != null)
						try {
							stmt2.close();
						} catch (SQLException e) {					
							e.printStackTrace();
						}	
				if(stmt_emp != null)
						try {
							stmt_emp.close();
						} catch (SQLException e) {					
							e.printStackTrace();
						}
		    }
		    return result;		
	}
	//半年總分排名計算 End

//季總分排名計算 Start	
	public int QuarterRanking(String dept_id,String section,Connection conn)throws Exception{
		int result=0;
		Statement stmt = null;	
		Statement stmt2 = null;
		Statement stmt_emp = null;
		ResultSet rs;		
		ResultSet rs2;
		ResultSet rs_emp;
		String sSQL="";		
		String emp_id="";
		String year="";
		String month="";
		String score="";
		String bonus="";
		String balance="";
		String ranking_shift="";
		String ranking_station="";
		String ranking_title="";
		String ranking_group1="";
		String ranking_groupa="";
		String ranking_shifta="";
		int insertcheck=0;	
		java.util.LinkedList sql_insert = new java.util.LinkedList();			
		try{
				stmt = conn.createStatement();			
				stmt2 = conn.createStatement();	
				 //Add by Sam  on 20130715 for ReqNo:JC201300204 ,人員名單從原本取得分數的來源抽出單獨取得,將人員資料傳入分開計算分數與獎金
				stmt_emp = conn.createStatement();	
				sSQL = "   select DISTINCT a1.emp_id,a1.year," +
						"   case when to_char(add_months(sysdate, -1), 'mm') <= '03' then 'Q1' " +
						"   when to_char(add_months(sysdate, -1), 'mm') <= '06' then 'Q2' " +
						"   when to_char(add_months(sysdate, -1), 'mm') <= '09' then 'Q3'           " +
						"   else 'Q4' end as month" +
						"   from rbl_dl_performance_summary a1 , " +
						"   (select distinct bb1.emp_id,bb1.title,bb1.station_id,bb1.shift_id,bb1.dept_id,bb1.position_group  " +
						"   from rbl_dl_emp bb1 ,rbl_dl_organization bb2  " +
						"   where bb1.dept_id = bb2.dept_id and bb1.shift_id = bb2.shift_id and bb1.station_id = bb2.station_id " +
						"   and bb1.dept_id ='"+dept_id+"' and bb2.section ='"+section+"' " +
						"   ) b1 " +
						"   where a1.emp_id = b1.emp_id and b1.dept_id ='"+dept_id+"'      " +
						"   and a1.year = to_char(add_months(sysdate, -1), 'yyyy')" +
						"   and a1.year || a1.month >= to_char(add_months(sysdate, -3), 'yyyymm')" +
					 	"   and a1.year || a1.month <= to_char(add_months(sysdate, -1), 'yyyymm')";
				stmt_emp.executeQuery(sSQL);				
				rs_emp = stmt_emp.getResultSet();
				while(rs_emp.next()){
					emp_id = rs_emp.getString("emp_id");
					year = rs_emp.getString("year");
					month = rs_emp.getString("month");
					//Modify by Sam on 20121203,JC201200344 ,增加判斷 rbl_dl_performance_summary.SkipFlag = 'N' ,Bonus及balance獨立運算				
					sSQL=" select emp_id , year,month ,title , score , ranking_shift,ranking_station,ranking_title,ranking_group1,ranking_groupa,ranking_shifta from  " +
					    "( select a.emp_id,a.year,b.title,"+
						"        case when to_char(add_months(sysdate, -1), 'mm') <= '03' then 'Q1' "+
						"			  when to_char(add_months(sysdate, -1), 'mm') <= '06' then 'Q2' "+
						"			  when to_char(add_months(sysdate, -1), 'mm') <= '09' then 'Q3' "+					
						"		 	  else 'Q4' end as month , a.score,"+
						"        decode(a.score,null,null,rank() over(partition by b.shift_id, b.title order by nvl(a.score, 0) desc)) as ranking_shift,"+
						"        decode(a.score,null,null,rank() over(partition by b.shift_id, b.title, b.station_id order by nvl(a.score, 0) desc)) as ranking_station,"+
						"        decode(a.score,null,null,rank() over(partition by b.title order by nvl(a.score, 0) desc)) as ranking_title,"+
						"        decode(a.score,null,null,rank() over(partition by b.title, b.position_group order by nvl(a.score, 0) desc)) as ranking_group1,"+
						"        decode(a.score,null,null,fun_get_ranking_level(a.emp_id, b.title, b.position_group,b.dept_id,a.year," +
						"        case when to_char(add_months(sysdate, -1), 'mm') <= '03' then 'Q1' "+
						"			  when to_char(add_months(sysdate, -1), 'mm') <= '06' then 'Q2' "+
						"			  when to_char(add_months(sysdate, -1), 'mm') <= '09' then 'Q3' "+					
						"		 	  else 'Q4' end,'"+section+"' )) as ranking_groupa,"+
						"        decode(a.score,null,null,fun_get_shift_ranking_level(a.emp_id, b.title, b.shift_id,b.dept_id,a.year,"+
						"		 case when to_char(add_months(sysdate, -1), 'mm') <= '03' then 'Q1' "+
						"			  when to_char(add_months(sysdate, -1), 'mm') <= '06' then 'Q2' "+
						"			  when to_char(add_months(sysdate, -1), 'mm') <= '09' then 'Q3' "+					
						"		 	  else 'Q4' end,'"+section+"' )) as ranking_shifta"+
						"   from (select a1.emp_id,a1.year,round(avg(score),2) as score,round(sum(bonus)) as bonus,round(avg(balance),2) as balance "+
						"           from rbl_dl_performance_summary a1 , "+
						"			(select distinct bb1.emp_id,bb1.title,bb1.station_id,bb1.shift_id,bb1.dept_id,bb1.position_group " + 
						"			 from rbl_dl_emp bb1 ,rbl_dl_organization bb2  " +
						"			 where bb1.dept_id = bb2.dept_id and bb1.shift_id = bb2.shift_id and bb1.station_id = bb2.station_id "+
						"			 and bb1.dept_id ='"+dept_id+"' and bb2.section ='"+section+"' "+
						"			 ) b1 "+
						"          where a1.emp_id = b1.emp_id and b1.dept_id ='"+dept_id+"'      "+
						"            and a1.year = to_char(add_months(sysdate, -1), 'yyyy')"+
						"            and a1.year || a1.month >= to_char(add_months(sysdate, -3), 'yyyymm')"+
						"            and a1.year || a1.month <= to_char(add_months(sysdate, -1), 'yyyymm')"+
						"            and a1.skipflag='N' "+
						"          group by a1.emp_id,a1.year ) a,  (select distinct emp_id,title,station_id,shift_id,dept_id,position_group from rbl_dl_emp ) b"+					
						"          where a.emp_id = b.emp_id )" +
						" where emp_id='"+emp_id+"' ";				
					stmt.executeQuery(sSQL);				
					rs = stmt.getResultSet();								
		            if(rs.next()){	            	
		            	//emp_id = rs.getString("emp_id");
		            	//year = rs.getString("year");
		            	//month = rs.getString("month");
		            	score = rs.getString("score");
		            	//bonus = rs.getString("bonus");
		            	//balance = rs.getString("balance");
		            	ranking_shift=rs.getString("ranking_shift");	            		
		            	ranking_station=rs.getString("ranking_station");
		            	ranking_title=rs.getString("ranking_title");
		            	ranking_group1=rs.getString("ranking_group1");
		            	ranking_groupa=rs.getString("ranking_groupa");			
		            	ranking_shifta=rs.getString("ranking_shifta");
		            }else{
		            	score ="";
		            	ranking_shift ="";
		            	ranking_station ="";
		            	ranking_title ="";
		            	ranking_group1 ="";
		            	ranking_groupa ="";
		            	ranking_shifta="";
		            }
	            	//Modify by Sam on 20121205,JC201200344 BONUS跟BALANCE獨立取得,不受SkipFlag限制
	            	sSQL=" select a.emp_id ,a.bonus,nvl(c.balance,0) as balance " +
            		" from (select a1.emp_id, round(sum(nvl(bonus,0))) as bonus  from rbl_dl_performance_summary a1 "+
            		" 			where a1.emp_id ='"+emp_id+"' "+ 
					"            and a1.year = to_char(add_months(sysdate, -1), 'yyyy')"+
					"            and a1.year || a1.month >= to_char(add_months(sysdate, -3), 'yyyymm')"+
					"            and a1.year || a1.month <= to_char(add_months(sysdate, -1), 'yyyymm')"+
					"		group by a1.emp_id ) a "+
					" ,(Select c1.emp_id,c1.balance from rbl_dl_performance_summary c1 where c1.year || c1.month = to_char(add_months(sysdate, -1), 'yyyymm')  ) c "+
					" where a.emp_id = c.emp_id(+) ";	            	
					stmt2.executeQuery(sSQL);				
					rs2 = stmt2.getResultSet();					
					if (rs2.next()){
						bonus = rs2.getString("bonus");
						balance = rs2.getString("balance");
					}else{
						bonus = "0";
						balance ="0";
					}	
					
            		sSQL="insert into Rbl_DL_Performance_summary(emp_id,year,month,score,bonus,balance,"+
            			 " ranking_shift,ranking_station,ranking_title,ranking_group1,ranking_groupa,ranking_shifta,"+
            			 " createuserid,createtime) values (?,?,?,?,?,?,?,?,?,?,?,?,"+
            			 "'SYS',to_char(sysdate,'yyyymmdd hh24miss')||'000')";
            		sql_insert.clear();
            		sql_insert.add(sSQL);
            		sql_insert.add(emp_id);
            		sql_insert.add(year);
            		sql_insert.add(month);
            		sql_insert.add(score);
            		sql_insert.add(bonus);
            		sql_insert.add(balance);     
            		sql_insert.add(ranking_shift);
            		sql_insert.add(ranking_station);
            		sql_insert.add(ranking_title);
            		sql_insert.add(ranking_group1);
            		sql_insert.add(ranking_groupa);
            		sql_insert.add(ranking_shifta);
            		insertcheck = insert(sql_insert,conn);	            				            	            			            		            
				}
					         	              
	            result = 1;
		    }catch(Exception e){		    	
		    	e.printStackTrace();
		    	result = 0;
		    }finally{
		    	if(stmt != null)
					try {
						stmt.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}	
				if(stmt2 != null)
					try {
						stmt2.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}	
				if(stmt_emp != null)
					try {
						stmt_emp.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}
		    }
		    return result;		
	}
	//季總分排名計算 End	
	
//年總分排名計算 Start
	
	public int TotalYearRanking(String dept_id,String section,Connection conn)throws Exception{
		int result=0;
		Statement stmt = null;	
		Statement stmt2 = null;
		Statement stmt_emp = null;
		ResultSet rs;
		ResultSet rs2;
		ResultSet rs_emp;
		String sSQL="";		
		String emp_id="";
		String year="";
		String month="";
		String score="";
		String bonus="";
		String balance="";
		String ranking_shift="";
		String ranking_station="";
		String ranking_title="";
		String ranking_group1="";
		String ranking_groupa="";
		String ranking_shifta="";
		int insertcheck=0;	
		java.util.LinkedList sql_insert = new java.util.LinkedList();			
		try{
				stmt = conn.createStatement();		
				stmt2 = conn.createStatement();		
				//Add by Sam  on 20130715 for ReqNo:JC201300204 ,人員名單從原本取得分數的來源抽出單獨取得,將人員資料傳入分開計算分數與獎金
				stmt_emp = conn.createStatement();	
				sSQL = "   select distinct  a1.emp_id,a1.year,'Y' AS month" +
                     "   from rbl_dl_performance_summary a1 , " +
                     "   (select distinct  bb1.emp_id,bb1.title,bb1.station_id,bb1.shift_id,bb1.dept_id,bb1.position_group " +
                     "   from  rbl_dl_emp bb1,rbl_dl_organization bb2 " +
                     "   where bb1.dept_id = bb2.dept_id and bb1.shift_id = bb2.shift_id and bb1.station_id = bb2.station_id  " +
                     "   and bb1.dept_id ='"+dept_id+"' and bb2.section ='"+section+"' " +
                     "   ) b1 " +
                     "   where a1.emp_id = b1.emp_id and b1.dept_id ='"+dept_id+"'      " +
                     "   and a1.year || a1.MONTH >= to_char(add_months(sysdate, -12), 'yyyy') || '01' " +
                     "   and a1.year || a1.MONTH <= to_char(add_months(sysdate, -12), 'yyyy') || '12' " +
                     "   ";
				stmt_emp.executeQuery(sSQL);				
				rs_emp = stmt_emp.getResultSet();
				while(rs_emp.next()){
					emp_id = rs_emp.getString("emp_id");
					year = rs_emp.getString("year");
					month = rs_emp.getString("month");
					//Modify by Sam on 20121203,JC201200344 ,增加判斷 rbl_dl_performance_summary.SkipFlag = 'N',Bonus及balance獨立運算				
					sSQL=" select emp_id , year,month ,title , score , ranking_shift,ranking_station,ranking_title,ranking_group1,ranking_groupa,ranking_shifta from " +
					" (select a.emp_id,a.year,b.title,"+
					"        'Y' as month , a.score,"+
					"        decode(a.score,null,null,rank() over(partition by b.shift_id, b.title order by nvl(a.score, 0) desc)) as ranking_shift,"+
					"        decode(a.score,null,null,rank() over(partition by b.shift_id, b.title, b.station_id order by nvl(a.score, 0) desc)) as ranking_station,"+
					"        decode(a.score,null,null,rank() over(partition by b.title order by nvl(a.score, 0) desc)) as ranking_title,"+
					"        decode(a.score,null,null,rank() over(partition by b.title, b.position_group order by nvl(a.score, 0) desc)) as ranking_group1,"+
					"        decode(a.score,null,null,fun_get_ranking_level(a.emp_id, b.title, b.position_group,b.dept_id,a.year,'Y','"+section+"' )) as ranking_groupa,"+
					"        decode(a.score,null,null,fun_get_shift_ranking_level(a.emp_id, b.title, b.shift_id,b.dept_id,a.year,'Y','"+section+"' )) as ranking_shifta"+
					"   from (select a1.emp_id,a1.year,round(avg(score),2) as score,round(sum(bonus)) as bonus,round(avg(balance),2) as balance "+
					"           from rbl_dl_performance_summary a1 , "+
					"			(select distinct  bb1.emp_id,bb1.title,bb1.station_id,bb1.shift_id,bb1.dept_id,bb1.position_group "+
					"			 from  rbl_dl_emp bb1,rbl_dl_organization bb2 "+
					"			where bb1.dept_id = bb2.dept_id and bb1.shift_id = bb2.shift_id and bb1.station_id = bb2.station_id  " +
					"			and bb1.dept_id ='"+dept_id+"' and bb2.section ='"+section+"' "+
					" 			) b1 "+
					"          where a1.emp_id = b1.emp_id and b1.dept_id ='"+dept_id+"'      "+
					"            and a1.year || a1.MONTH >= to_char(add_months(sysdate, -12), 'yyyy') || '01' "+
					"            and a1.year || a1.MONTH <= to_char(add_months(sysdate, -12), 'yyyy') || '12' "+
					"            and a1.skipflag='N' "+
					"          group by a1.emp_id,a1.year ) a,  (select distinct emp_id,title,station_id,shift_id,dept_id,position_group from rbl_dl_emp ) b"+					
					"          where a.emp_id = b.emp_id ) " +
					" where emp_id='"+emp_id+"' ";					
					stmt.executeQuery(sSQL);				
					rs = stmt.getResultSet();								
		            if(rs.next()){	            	
		            	//emp_id = rs.getString("emp_id");
		            	//year = rs.getString("year");
		            	//month = rs.getString("month");
		            	score = rs.getString("score");
		            	//bonus = rs.getString("bonus");
		            	//balance = rs.getString("balance");
		            	ranking_shift=rs.getString("ranking_shift");	            		
		            	ranking_station=rs.getString("ranking_station");
		            	ranking_title=rs.getString("ranking_title");
		            	ranking_group1=rs.getString("ranking_group1");
		            	ranking_groupa=rs.getString("ranking_groupa");			
		            	ranking_shifta=rs.getString("ranking_shifta");
		            }else{
		            	score ="";
		            	ranking_shift ="";
		            	ranking_station ="";
		            	ranking_title ="";
		            	ranking_group1 ="";
		            	ranking_groupa ="";
		            	ranking_shifta="";
		            }
	            	//Modify by Sam on 20121205,JC201200344 BONUS跟BALANCE獨立取得,不受SkipFlag限制
	            	sSQL=" select a.emp_id ,a.bonus,nvl(c.balance,0) as balance " +
            		" from (select a1.emp_id, round(sum(nvl(bonus,0))) as bonus from rbl_dl_performance_summary a1 "+
            		" 			where a1.emp_id ='"+emp_id+"' "+ 					
					"            and a1.year || a1.MONTH >= to_char(add_months(sysdate, -12), 'yyyy') || '01' "+
					"            and a1.year || a1.MONTH <= to_char(add_months(sysdate, -12), 'yyyy') || '12' "+
					"		group by a1.emp_id ) a "+
					"		 ,(Select c1.emp_id,c1.balance from rbl_dl_performance_summary c1 where c1.year || c1.month = to_char(add_months(sysdate, -12), 'yyyy') || '12'  ) c "+
					" where a.emp_id = c.emp_id(+) ";	            	
					stmt2.executeQuery(sSQL);				
					rs2 = stmt2.getResultSet();					
					if (rs2.next()){
						bonus = rs2.getString("bonus");
						balance = rs2.getString("balance");
					}else{
						bonus = "0";
						balance ="0";
					}	
					
            		sSQL="insert into Rbl_DL_Performance_summary(emp_id,year,month,score,bonus,balance,"+
            			 " ranking_shift,ranking_station,ranking_title,ranking_group1,ranking_groupa,ranking_shifta,"+
            			 " createuserid,createtime) values (?,?,?,?,?,?,?,?,?,?,?,?,"+
            			 "'SYS',to_char(sysdate,'yyyymmdd hh24miss')||'000')";
            		sql_insert.clear();
            		sql_insert.add(sSQL);
            		sql_insert.add(emp_id);
            		sql_insert.add(year);
            		sql_insert.add(month);
            		sql_insert.add(score);
            		sql_insert.add(bonus);
            		sql_insert.add(balance);     
            		sql_insert.add(ranking_shift);
            		sql_insert.add(ranking_station);
            		sql_insert.add(ranking_title);
            		sql_insert.add(ranking_group1);
            		sql_insert.add(ranking_groupa);
            		sql_insert.add(ranking_shifta);
            		insertcheck = insert(sql_insert,conn);	            		
		            	            			            		            
				}
					         	              
	            result = 1;
		    }catch(Exception e){		    	
		    	e.printStackTrace();
		    	result = 0;
		    }finally{
		    	if(stmt != null)
					try {
						stmt.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}		   
				if(stmt2 != null)
					try {
						stmt2.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}						
				if(stmt_emp != null)
					try {
						stmt_emp.close();
					} catch (SQLException e) {					
						e.printStackTrace();
					}					
		    }
		    return result;		
	}
	//年總分排名計算 End	
//	更新狀態 Start
	public int updateStatus(String dept_id,String section ,Connection conn){
		int result=0;		
		String sSQL="";
		java.util.LinkedList sql_update = new java.util.LinkedList();
		int updatecheck=0;
		try{
			//Modify by Sam on 20130311,增加section條件
			sSQL="update Rbl_dl_status set status=? ,updatetime =to_char(sysdate,'yyyymmdd hh24miss') || '000',updateuserid='SYS' where dept_id=? and type='SUBJECT' and year || month=to_char(add_months(sysdate, -1), 'YYYYMM') " +
			     " and  (dept_id,shift_id,station_id) in (select dept_id,shift_id,station_id from rbl_dl_organization where section ='"+section+"' )  ";
    		sql_update.clear();	            	
    		sql_update.add(sSQL);  
    		//sql_update.add("CALCULATE_2 OK");   Modify by Sam on 20130322,統一改成SUMMARY OK
    		sql_update.add("SUMMARY OK");
    		sql_update.add(dept_id);
    		updatecheck = update(sql_update,conn);	    			
			result = 1;	
		}catch(Exception e){		    	
	    	e.printStackTrace();
	    	result = 0;
	    }
	    return result;	
	}
	
//	更新狀態 End	
	
//	 insert(LinkedList list,Connection conn)
	  public int insert(LinkedList list,Connection conn) {
	    int check_list = 0;
	    PreparedStatement ps;
	    try {	     
	      //System.out.println((String)list.get(0));
	      ps= conn.prepareStatement( (String)list.get(0) );
	      for (int i=1; i<list.size(); i++){
	      //System.out.println(i);
	        if ((list.get(i)) instanceof java.sql.Date)
	          ps.setDate( i,(java.sql.Date)list.get(i) );
	        else if ((list.get(i)) instanceof java.sql.Timestamp)
	       	  ps.setTimestamp( i,(java.sql.Timestamp)list.get(i) );
	        else if(list.get(i) == null)
	          ps.setNull( i , Types.VARCHAR);
	        else
	          ps.setString( i,((String)list.get(i)));
	      }
	      check_list = ps.executeUpdate();	      
	    }
	    catch( Exception e ) {
	      e.printStackTrace();	      
	    }
	    return check_list;
	  }// end of insert(LinkedList list,Connection conn)
//	start of update(LinkedList list,Connection conn)
	  public int update(LinkedList list,Connection conn) {
	      int check_list = 0;
	      PreparedStatement ps;
	      try {     
	        ps= conn.prepareStatement( (String)list.get(0) );
	        for (int i=1; i<list.size(); i++){
	          if ((list.get(i)) instanceof java.sql.Date)
	            ps.setDate( i,(java.sql.Date)list.get(i) );          
	          else if ((list.get(i)) instanceof java.sql.Timestamp)
	         	  ps.setTimestamp( i,(java.sql.Timestamp)list.get(i) );
	          else if(list.get(i) == null)
	            ps.setNull( i , Types.VARCHAR);      
	          else
	          	ps.setString( i,((String)list.get(i)));	           
	        }      
	        check_list = ps.executeUpdate();      
	       }
	      catch( Exception e ) {
	        e.printStackTrace();      
	      }
	      return check_list;
	    } // end of update(LinkedList list,Connection conn)

	  public ResultSet selectData( String str,Connection conn ) {
		    ResultSet result = null;
		    Statement dps = null;
		    try {
		      dps = conn.createStatement();
		      result = dps.executeQuery( str );			      
		      dps.close();
		    }
		    catch( Exception e ) {		    
		      e.printStackTrace();
		    }		    
		    return result;
		  }

	public static void main(String[] args) {
	        Connection conn = null;
	        String sqlstmt;	
	        Statement stmt = null;
	        Statement stmt2 = null;
	        ResultSet rs;		        
	        ResultSet rs2;
	        String dept_id ="";
	        String year="";
	        String month="";
	        String section = "";
	        int rowcount=0;	       
	        int result =0;
	        boolean bFind = true;
	        try{
	        	System.out.println("Get Connection");
	            conn = DBConn.getOracleConnection();
	            DBConn.setConnectionAutoCommit(conn,false);
	            stmt = conn.createStatement();
	            stmt2 = conn.createStatement();
	          //Modify by Sam on 20130308 ,增加section ,PROD專用	           
	            sqlstmt="select dept_id,section , count(*) as cnt from ( "+
		        		" select distinct t.dept_id,t.status,b.section  from dlowner.rbl_dl_status t,rbl_dl_organization b  "+
		        		"	where t.year || t.month = to_char(add_months(sysdate, -1), 'YYYYMM') "+
		        		"   and t.dept_id =b.dept_id and t.shift_id =b.shift_id and t.station_id = b.station_id "+
		        		"	AND t.TYPE = 'SUBJECT' and b.section = 'PROD' ) "+
		        		"	group by dept_id,section "; 
	            stmt.executeQuery(sqlstmt);	            
	            rs = stmt.getResultSet();		            	            
	            
	            Month_Cal_2 mode1 = new Month_Cal_2();	            
	            while(rs.next()){
	            	dept_id = rs.getString("dept_id");
	            	section = rs.getString("section");
	            	System.out.println("Dept_ID:"+dept_id+" Section:"+section+ " Process");
	            	if(rs.getInt("cnt")==1){
	            		//Modify by Sam on 20130308 ,增加section	            		
	            		sqlstmt = "select distinct t.status,t.month,t.year from dlowner.rbl_dl_status t,rbl_dl_organization b "+
			      				  "	where t.year || t.month = to_char(add_months(sysdate,-1 ), 'YYYYMM') " +
			      				  " and t.dept_id = b.dept_id and t.shift_id = b.shift_id and t.station_id = b.station_id " +
			      				  " AND t.TYPE='SUBJECT' AND t.Dept_ID='"+dept_id+"' and b.section ='"+section+"' ";
	            		stmt2.executeQuery(sqlstmt);	            
	    	            rs2 = stmt2.getResultSet();
	    	            if(rs2.next()){
	    	            	year=rs2.getString("year");
	    	            	month=rs2.getString("month");	    	            	
	    	            	 
	    	            	if(rs2.getString("status").equals("MANAGER OK")){
	    	            		result = mode1.getManagerDatatoSummary(dept_id,section,conn);
	    	            		if(result==1) System.out.println("getManagerDatatoSummary OK");
	    	            		else System.out.println("getManagerDatatoSummary Fail");
	    	            		result = mode1.getSummaryRanking(dept_id,section,conn);
	    	            		if(result==1) System.out.println("getSummaryRanking OK");
	    	            		else System.out.println("getSummaryRanking Fail");
	    	            		//季結算		Add by Sam on 20120312 
	    	            		if(month.equals("03")||month.equals("06")||month.equals("09")||month.equals("12")){
	    	            			result = mode1.SetSummarySkipFlag(dept_id,section,conn); //Add by Sam on 20121203 JC201200344 , update skipflag
	    	            			if(result==1) System.out.println("SetSummarySkipFlag OK");
	    	            			else System.out.println("SetSummarySkipFlag Fail"); 	  	
	    	            			
	       	            			result = mode1.QuarterRanking(dept_id,section,conn);	    	            		
	    	            			if(result==1) System.out.println("QuarterRanking OK");
	    	            			else System.out.println("QuarterRanking Fail"); 	  	            			
	    	            		}
	    	            		//半年結算
	    	            		if(month.equals("06")||month.equals("12")){ 
	    	            			result = mode1.sixMonthRanking(dept_id,section,conn);
	    	            			if(result==1) System.out.println("sixMonthRanking OK");
	    	            			else System.out.println("sixMonthRanking Fail");
	    	            		}
	    	            		//年結算 	Add by Sam on 20120312
	    	            		if(month.equals("12")){ 
	       	            			result = mode1.TotalYearRanking(dept_id,section,conn);	    	            		
	    	            			if(result==1) System.out.println("TotalYearRanking OK");
	    	            			else System.out.println("TotalYearRanking Fail"); 	   	            			
	    	            		}
	    	            		result = mode1.updateStatus(dept_id,section,conn);
	    	            		if(result==1) System.out.println("updateStatus OK");
	    	            		else System.out.println("updateStatus Fail");
	    	            	}
	    	            }
	            	}	            		
	            } 
	            
	        }catch(Exception ex){
	            System.out.println(ex);
	            DBConn.rollback(conn);
	        }finally{	        		        	
	        	DBConn.commit(conn);
	            DBConn.close(conn);
	            System.out.println("Process End");	            
	        }
	}	
}


根據您提供的 Java 程式碼，Rbl_DL_Performance_summary 資料表中的 score (分數) 欄位，會在 四種不同的結算情境 下被計算並 Insert。這四種情境分別是：單月結算、半年結算、季度結算 與 年度結算。

以下為您詳細拆解各個情境的計算來源 (Tables) 與條件 (Conditions)：

1. 單月結算 (getManagerDatatoSummary 方法)
這個方法負責抓取上個月的原始分數，並 Insert 到 summary 表中。

資料來源 Tables：

rbl_dl_performance_manager (經理評分資料表)

rbl_dl_emp (員工基本資料表)

rbl_dl_organization (組織單位資料表)

計算邏輯與條件：

鎖定人員與時間：先找出符合特定部門 (dept_id)、特定課別 (section)，且時間點為 系統日期的「上個月」 (to_char(add_months(sysdate,-1 ), 'YYYYMM')) 的員工。這部分透過 JOIN rbl_dl_emp 與 rbl_dl_organization 來確認人員隸屬的 shift_id 與 station_id。

擷取 Score (原始分數)：針對上述找出的員工 (emp_id)，去 rbl_dl_performance_manager 查詢其分數。

條件：year || month 為上個月。

條件：item like '經理評比%Ranking' (只抓取項目名稱為「經理評比...Ranking」的紀錄)。

計算：直接取該筆紀錄的 final 欄位值 (nvl(a.final,0))。如果查無資料，則 score 設為 "0"。

2. 半年結算 (sixMonthRanking 方法)
這個方法負責計算每半年的平均分數 (1H 或是 2H)，並將這筆半年總結紀錄 Insert 回 summary 表。

資料來源 Tables：

rbl_dl_performance_summary (自身的歷史紀錄)

rbl_dl_emp

rbl_dl_organization

計算邏輯與條件：

擷取 Score (半年平均分數)：

條件時間區間：系統日期的 過去 6 個月 內 (>= to_char(add_months(sysdate, -6), 'yyyymm') 且 <= to_char(add_months(sysdate, -1), 'yyyymm'))。

排除條件 (重要)：skipflag = 'N'。這表示只計算沒有被標記為 Skip 的月份。

計算公式：利用 round(avg(score), 2)，將該員工過去 6 個月內符合條件的 score 進行平均，並四捨五入至小數點後第二位。

3. 季度結算 (QuarterRanking 方法)
計算每季 (Q1, Q2, Q3, Q4) 的平均分數，並 Insert 回 summary 表。

資料來源 Tables：

rbl_dl_performance_summary

rbl_dl_emp

rbl_dl_organization

計算邏輯與條件：

擷取 Score (季度平均分數)：

條件時間區間：系統日期的 過去 3 個月 內 (>= to_char(add_months(sysdate, -3), 'yyyymm') 且 <= to_char(add_months(sysdate, -1), 'yyyymm'))。

排除條件：skipflag = 'N' (同樣排除不列入計算的月份)。

計算公式：round(avg(score), 2)，將該季內 (過去 3 個月) 的分數平均，取到小數點後第二位。

4. 年度結算 (TotalYearRanking 方法)
計算一整年度 (Y) 的平均分數，並 Insert 回 summary 表。

資料來源 Tables：

rbl_dl_performance_summary

rbl_dl_emp

rbl_dl_organization

計算邏輯與條件：

擷取 Score (年度平均分數)：

條件時間區間：計算對象年份的 1月到12月 (>= ... 'yyyy' || '01' 且 <= ... 'yyyy' || '12')。

排除條件：skipflag = 'N'。

計算公式：round(avg(score), 2)，將該年度內 12 個月的分數進行平均，取到小數點後第二位。

總結
最源頭的單月 score 是從 rbl_dl_performance_manager (經理評估表) 的 final 欄位抓過來的 (前提是 item 包含 經理評比%Ranking)。

季、半年、年度的 score 則都是去撈取 Rbl_DL_Performance_summary 過去自己產生的單月資料，加上條件 skipflag = 'N' 之後，使用 AVG() 函數做算術平均得來的。
