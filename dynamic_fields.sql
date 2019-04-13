select dfv.value_text, dfv.value_date, dfv.value_int,
       df.name, df.label, df.field_type
from dynamic_field_value as dfv
inner join dynamic_field as df on dfv.field_id = df.id
where dfv.object_id = {0}