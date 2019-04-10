select value_text, tt.name, count(*) as frequency
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
where field_id = 14
group by value_text, name;