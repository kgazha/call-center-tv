select field_id, value_text, t.id as ticket_id,
       t.tn, t.create_time, closed
from dynamic_field_value as dfv
inner join ticket as t on dfv.object_id = t.id
inner join ticket_type as tt on t.type_id = tt.id
left join (
    SELECT tid, create_time closed
    FROM ticket_history th
    INNER JOIN (SELECT MIN(id) id, ticket_id tid
                FROM ticket_history th
                WHERE
                    th.create_time > '{0}'
                GROUP BY tid) thids
    ON th.id = thids.id
    WHERE th.create_time > '{0}'
) s ON t.id = s.tid
where field_id in (14, 44)
and tt.id = 11
and ticket_state_id not in (5, 6, 9)
and t.create_time > '{0}'
order by t.id
