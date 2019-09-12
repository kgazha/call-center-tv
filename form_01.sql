select s1.value_text, name, ticket_type_id, frequency, complaints from
(
	select value_text, tt.name, tt.id as ticket_type_id, count(*) as frequency, field_id, t.id as ticket_id
	from dynamic_field_value as dfv
	inner join ticket as t on dfv.object_id = t.id
	inner join ticket_type as tt on t.type_id = tt.id
	inner join ticket_state as ts on ts.id = t.ticket_state_id
	where field_id = 14
	and tt.id > 7
	and ts.id not in (5, 6, 9)
	and t.create_time > '{0}'
	and t.create_time < '{1}'
	group by value_text, name
	order by value_text
) s1
left join
(
	select value_text, sum(complaints) as complaints
	from
	(
		select value_text, sum(value_int) as complaints, object_id
		from dynamic_field_value as dfv
		inner join ticket as t on dfv.object_id = t.id
		inner join ticket_type as tt on t.type_id = tt.id
		where field_id in (14, 37)
		and tt.id > 7
		and t.create_time > '{0}'
		group by object_id
	) s2
	group by value_text
) s3
on s1.value_text = s3.value_text;