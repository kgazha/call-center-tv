select count(*) as _count from ticket_history
where id in (
	select max(id) from
	(
		select * from ticket_history
		where create_time < '{0}'
	) s
	group by ticket_id
)
and state_id = 4;