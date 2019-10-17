select field_id, value_text, t.id as ticket_id, t.create_time, last_action_time,
       s.state_id as ticket_state_id, value_int, t.tn
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
right join (
	SELECT id, ticket_id as tid, create_time as last_action_time, th.state_id
	from ticket_history th
	where id in (
		select max(id) from
		(
			select * from ticket_history
			where create_time > '{0}'
            and create_time < '{1}'
		) s_id
		group by ticket_id
	)
) s ON t.id = s.tid
where field_id in (14, 12, 15, 17, 16, 37, 39, 40)
and tt.id = 11
and ticket_state_id not in (5, 6, 9)
order by value_int;