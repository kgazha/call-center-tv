select s1.field_id, s1.value_text, s1.ticket_id, s1.create_time, s1.ticket_state_id, s2.close_time, value_int from
(select field_id, value_text, t.id as ticket_id, t.create_time, ts.id as ticket_state_id, value_int
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
inner join ticket_state as ts on ts.id = t.ticket_state_id
where field_id in (14, 12, 15, 17, 16, 37)
and tt.id = 11
and ts.id in (2, 3, 10)
and t.create_time > '{0}'
and t.create_time <= '{1}'
order by t.id) s1
inner join
(select ticket_id, max(create_time) as close_time from ticket_history
group by ticket_id) s2
ON s1.ticket_id = s2.ticket_id;