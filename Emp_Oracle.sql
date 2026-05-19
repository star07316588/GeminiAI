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

WITH EMP_ITEM_BASE AS (
    -- 步驟 1：建立「應評比員工」與「對應評比項目」的完整基準清單
    SELECT e.emp_id, 
           e.dept_id, 
           e.title, 
           e.station_id, 
           e.shift_id,
           c.item, 
           d.detailitem
      FROM rbl_dl_emp e
      -- 關聯該職稱與站別對應的評比主項目 (Item)
      JOIN rbl_dl_item c 
        ON e.title = c.title 
       AND e.station_id = c.station_id
      -- 關聯評比細項 (DetailItem)
      JOIN rbl_dl_detailitem d 
        ON c.title = d.title 
       AND c.station_id = d.station_id 
       AND c.item = d.item
     WHERE e.dept_id = 'MK310'
       AND e.station_id = 'W'
       AND e.title = 'TA'
       AND e.shift_id IN ('C2', 'B1', 'C1', 'B2')
       AND c.subjective = 'Y'
       AND c.deleteflag = 'N'
       AND d.deleteflag = 'N'
)
-- 步驟 2：將基準清單與 2026 年 04 月的實際評比紀錄做 LEFT JOIN
SELECT base.emp_id,
       base.item,
       base.detailitem,
       p.record,
       p.comments,
       -- 建立一個狀態欄位方便網頁前端判斷或顯示
       CASE WHEN p.record IS NULL THEN '尚未評比' ELSE '已評比' END AS eval_status
  FROM EMP_ITEM_BASE base
  LEFT JOIN rbl_dl_performance_subject p
         ON base.emp_id = p.emp_id
        AND base.item = p.item
        AND base.detailitem = p.detailitem
        AND p.year = '2026'
        AND p.month = '04'
 ORDER BY base.emp_id, base.item, base.detailitem;
