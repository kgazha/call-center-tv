select * from
(select value_text, tt.name, tt.id as ticket_type_id, count(*) as frequency, field_id, t.id as ticket_id
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
where field_id = 14
and tt.id > 7
and t.create_time > '{0}'
group by value_text, name
order by value_text
) s1
left join
(select object_id, value_int, field_id as complaint_field_id
from dynamic_field_value as dfv
where field_id = 37
) s2
on s1.ticket_id = s2.object_id;