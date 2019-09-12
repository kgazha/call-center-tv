select field_id, value_text, t.id as ticket_id, t.create_time, closed,
       ts.id as ticket_state_id, value_int, t.tn
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
inner join ticket_state as ts on ts.id = t.ticket_state_id
left join (
    SELECT tid, create_time closed
    FROM ticket_history th
    INNER JOIN (SELECT MIN(id) id, ticket_id tid
                FROM ticket_history th
                WHERE
                    th.create_time > '{0}'
                    AND th.state_id in (2, 3, 10)
                GROUP BY tid) thids
    ON th.id = thids.id
    WHERE th.create_time > '{0}'
) s ON t.id = s.tid
where field_id in (14, 12, 15, 17, 16, 37, 39, 40)
and tt.id = 11
and ticket_state_id not in (5, 6, 9)
and t.create_time > '{0}'
and t.create_time < '{1}'
order by t.id
