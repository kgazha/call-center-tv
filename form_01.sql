select value_text, tt.name, tt.id as ticket_type_id, count(*) as frequency
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
where field_id = 14
and t.create_time > '{0}'
and tt.id > 7
group by value_text, name
order by value_text;