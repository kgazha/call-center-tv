select * from ticket_history
where id in (
	select max(id) from
	(
		select * from ticket_history
		where create_time > '{0}'
		and create_time < '{1}'
	) s
	group by ticket_id
);