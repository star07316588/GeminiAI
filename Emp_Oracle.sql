select * from rbl_dl_emp where dept_id='MK310' and station_id='W' and title='TA' and shift_id in ('C2', 'B1', 'C1', 'B2') order by emp_id

select distinct a.emp_id, a.item, a.detailitem, a.record, a.comments
  from rbl_dl_performance_subject a,
       rbl_dl_emp                 b,
       rbl_dl_item                c,
       rbl_dl_detailitem          d
 where a.year = '2026'
   and a.month = '04'
   and a.emp_id = b.emp_id(+)
   and b.title = 'TA'
   and b.station_id = 'W'
   and b.dept_id = 'MK310'
   and b.shift_id in ('C2', 'B1', 'C1', 'B2')
   and b.title = c.title(+)
   and b.station_id = c.station_id(+)
   and c.station_id = d.station_id(+)
   and c.title = d.title(+)
   and c.item = d.item(+)
   and c.subjective(+) = 'Y'
   and c.deleteflag(+) = 'N'
   and d.deleteflag(+) = 'N'
 order by a.emp_id, a.item, a.detailitem
