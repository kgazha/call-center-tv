select field_id, value_text, t.id as ticket_id, t.create_time, value_int
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
where field_id in (14, 12, 15, 17, 16, 18, 37)
and t.create_time > '{0}'
and t.create_time <= '{1}'
and tt.id = 10
order by t.id;