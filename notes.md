# My Notes

``` sql
create table lots as SELECT cast(locationid as NUMERIC) as id, cast(latitude as NUMERIC) as lat, cast(longitude as NUMERIC) as lon, address1 as address, city as city, state as state, locationName as name, zip5 as zip5, zip4 as zip4, locationType as type, 0 as posted from tmp2
```

``` sql
select offices2.*
from offices2
where offices2.locationID not in
(
select id from lots
);
```
