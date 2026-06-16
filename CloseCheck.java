	public static int CloseCheck(Connection con,String sType,String logintitle,
								 String year,String month,
								 String station_id,String title,
								 String dept_id,String Shift_id) throws Exception {
		
		int Result =1;
		Statement stmt = null;
		Statement stmt2 = null;
		Statement stmt3 = null;
		String sSQL ="";
		ResultSet rs;
		ResultSet rs2;
		ResultSet rs3;
		int icnt =0;	
		int peoples =0;
		boolean bbreak = false;
		station_id = station_id.replace('(',' ');
		station_id = station_id.replace(')',' ');
		title = title.replace('(',' ');
		title = title.replace(')',' ');
		Shift_id = Shift_id.replace(')',' ');
		Shift_id = Shift_id.replace('(',' ');
		String[] stationArray=station_id.split(",");
		String[] titleArray=title.split(",");
		String[] shiftArray=Shift_id.split(",");
		String sTable="";
		if(sType.equals("SUBJECT"))
			sTable="dlowner.rbl_dl_performance_subject";
		else
			sTable="dlowner.rbl_dl_performance_manager";		
		try{
		stmt = con.createStatement();		
		stmt2 = con.createStatement();		
		stmt3 = con.createStatement();		
		sSQL="select distinct a.item, a.detailitem"+
		  	" from rbl_dl_detailitem a, rbl_dl_item b"+
		  	" where a.station_id = b.station_id"+
		    " and a.title = b.title"+
		    " and a.item = b.item"+
		    " and a.DELETEFLAG ='N'"+
		    " and b.DELETEFLAG ='N'"+
		    " and b.subjective = 'Y'"+
		    " and a.station_id in ("+station_id+") "+
		    " and a.title in ("+title+")"+
		    " and (instr(input,'"+logintitle+"') <> 0 or input = 'ALL') ";
		System.out.println(sSQL);
		stmt.executeQuery(sSQL);				
		rs = stmt.getResultSet();						
        while(rs.next()){        	
        	for(int x=0 ;x<stationArray.length;x++){
        		for(int y=0 ;y<titleArray.length;y++){
        			for(int z=0 ;z<shiftArray.length;z++){
        				peoples =0;
        				sSQL = " select count(*) as peoples "+
 				  	   		   " from sbl_emp b "+
 				  	   		   " where b.title ="+titleArray[y]+" and b.station_id ="+stationArray[x]+" "+
 				  	   		   " and b.shift_id ="+shiftArray[z]+" and b.dept_id="+dept_id+" AND B.DLT_USER IS NULL ";
    					stmt3.executeQuery(sSQL);				
        				rs3 = stmt3.getResultSet();
        				while(rs3.next()){
        					peoples = rs3.getInt("peoples");
        				}	
        				sSQL = " select count(*) as cnt "+
        				  	   " from "+sTable+" a, sbl_emp b "+
        				  	   " where a.emp_id = b.emp_id   "+
        				  	   " and a.month = '"+month+"' and a.year = '"+year+"' "+
        				  	   " and a.item = '"+rs.getString("item")+"' and a.detailitem = '"+rs.getString("detailitem")+"' "+
        				  	   " and b.title ="+titleArray[y]+" and b.station_id ="+stationArray[x]+" "+
        				  	   " and b.shift_id ="+shiftArray[z]+" and b.dept_id="+dept_id+" ";
        				if(sType.equals("MANAGER"))
        					sSQL+=" and a.final is not null ";
        				//System.out.println(sSQL);
    					stmt2.executeQuery(sSQL);				
        				rs2 = stmt2.getResultSet();	
        				while(rs2.next())
        					icnt = rs2.getInt("cnt");        					
        					if(icnt==0 && peoples >0){
        						bbreak=true;
        						Result =0;
        					}
        				if(bbreak==true) break;
        			}
        			if(bbreak==true) break;
        		}
        		if(bbreak==true) break;
        	}
        	if(bbreak==true) break;
        }
		
    }catch(Exception e){		    	
    	e.printStackTrace();   
    	Result =0;
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
		if(stmt3 != null)
			try {
				stmt3.close();
			}catch (SQLException e) {					
				e.printStackTrace();
			}
    }
	return Result;			
	}
