select s1.field_id, s1.value_text, s1.ticket_id, s1.create_time, a_body, value_int
from
(select field_id, value_text, t.id as ticket_id, t.create_time, value_int
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
where field_id in (14, 12, 15, 37)
and tt.id = 14
and t.create_time > '{0}'
order by t.id) s1
inner join
(select ticket_id, max(id) as article_id
from article
group by ticket_id) s2
ON s1.ticket_id = s2.ticket_id
inner join article_data_mime as adm on s2.article_id = adm.article_id
