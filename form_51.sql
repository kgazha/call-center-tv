select field_id, value_text, t.id as ticket_id, t.create_time, ts.id as ticket_state_id
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
inner join ticket_state as ts on ts.id = t.ticket_state_id
where field_id in (14, 30)
and tt.id = 11
order by t.id;