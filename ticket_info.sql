select t.id, t.tn, t.title, t.create_time,
       q.name as queue, tt.name as ticket_type
from ticket as t
inner join ticket_type as tt on t.type_id = tt.id
inner join queue as q on t.queue_id = q.id;