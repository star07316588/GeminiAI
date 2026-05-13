	    var year = document.getElementById("year").value;	    
	    var month = document.getElementById("month").value;		
		var title = document.getElementById("title").value;						
		var dept_id = document.getElementById("dept_id").value;		
		var station_id = document.getElementById("station_id").value;
		var action = document.getElementById("actionevent").value;		
		var shift_id = document.getElementById("shift_id");

<tbody>
	<tr class="searchArea">
		<c:if test="${actionevent!='Query'}">
			<td>			
			<select name="year">		
					<c:forEach items="${years}" var = "item" varStatus="st">		
	                	<option value="<c:out value="${item.YEAR}"/>" <c:if test="${item.YEAR==year}">selected</c:if>><c:out value="${item.YEAR}"/></option>
	                	<c:if test="${item.YEAR != item.LASTYEAR}">
	                		<option value="<c:out value="${item.LASTYEAR}"/>" <c:if test="${item.LASTYEAR==year }">selected</c:if>><c:out value="${item.LASTYEAR}"/></option>
	                	</c:if>
	                </c:forEach>                
			</select>						
			</td>
			<td>	
			<select name="month">		
					<c:forEach items="${months}" var = "item" varStatus="st">				
	                	<option value="<c:out value="${item.MONTH}"/>" <c:if test="${item.MONTH==month }">selected</c:if>><c:out value="${item.MONTH}"/></option>                	
	                	<option value="<c:out value="${item.LASTMONTH}"/>" <c:if test="${item.LASTMONTH==month }">selected</c:if>><c:out value="${item.LASTMONTH}"/></option>                	
	                </c:forEach>                
			</select>
			</td>
			<td>
				<input type="hidden" name ="dept_id" value="<c:out value="${loginUser.DEPT_ID}"/>"/>
	           	<c:out value="${loginUser.DEPT_ID}"/>
			</td>
			<td>
				<input type="hidden" name ="station_id" value="<c:out value="${loginUser.STATION_ID}"/>"/>
	           	<c:out value="${loginUser.STATION_ID}"/>
			</td>
			<td>	
			<select name="title">
					<c:forEach items="${titles}" var = "item" varStatus="st">
	                	<option value="<c:out value="${item}"/>" <c:if test="${item==title }">selected</c:if>><c:out value="${item}"/></option>
	                </c:forEach>
			</select>
			</td>	
			<td>	
			<select  id="shift_id" name="shift_id" multiple="multiple" >
					<c:forEach items="${shift_ids}" var = "item" varStatus="st">
						<c:set var="shift"  value="N"/>
						<c:forEach items="${shift_id}" var = "item2" varStatus="st">
	                		<c:if test="${item==item2 }">
	                			<c:set var="shift"  value="${item2}"/>	
	                		</c:if>	                		
	                	</c:forEach>
	                	<option value="<c:out value="${item}"/>" <c:if test="${item==shift}">selected</c:if>><c:out value="${item}"/></option>
	                </c:forEach>
			</select>
			</td>	
			<td>
			<input type="image" name="Image5" src="../images/confirm.gif"  >
			</td>
		</c:if>
		<c:if test="${actionevent=='Query'}">
			<td><input type="hidden" name ="year" value="<c:out value="${year}"/>"/><c:out value="${year}"></c:out></td>
			<td><input type="hidden" name ="month" value="<c:out value="${month}"/>"/><c:out value="${month}"></c:out></td>
			<td><input type="hidden" name ="dept_id" value="<c:out value="${loginUser.DEPT_ID}"/>"/><c:out value="${loginUser.DEPT_ID}"></c:out></td>
			<td><input type="hidden" name ="station_id" value="<c:out value="${loginUser.STATION_ID}"/>"/><c:out value="${loginUser.STATION_ID}"></c:out></td>
			<td><input type="hidden" name ="title" value="<c:out value="${title}"/>"/><c:out value="${title}"></c:out></td>
			<td>
				<c:set var="shift"  value="N"/>
				<c:forEach items="${shift_id}" var = "item" varStatus="st">
					<c:if test="${shift!='N'}">
						<c:set var="shift"  value="${shift},${item}"/> 
					</c:if>					   					
					<c:if test="${shift=='N'}">
						<c:set var="shift"  value="${item}"/> 
					</c:if>	        					     	        	    		
	        	</c:forEach>	        	
	        	<input type="hidden" name ="shift_id" value="<c:out value="${shift}"/>"/><c:out value="${shift}"></c:out>	        	
			</td>
			<td>
			<input type="image" name="Image5" src="../images/confirm.gif"  >
			</td>
		</c:if>		
	</tr>
</tbody>
</table>
