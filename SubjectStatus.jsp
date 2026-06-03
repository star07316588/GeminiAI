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
