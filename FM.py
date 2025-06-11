from datetime import datetime, timedelta

base_url = (
    "https://reboks.nus.edu.sg/nus_public_web/public/index.php/facilities/group_booking"
    "?csrf_test_name=a594ac8e4f45790f2f8b7632fd95d58c"
    "&group_filter_one=94"
    "&group_filter_two=108"
    "&group_activity_filter=267"
    "&group_venue_filter=15"
    "&group_subvenue_filter="
    "&day_filter="
    "&date_filter_from={date}"
    "&date_filter_to={date}"
    "&time_filter_from="
    "&time_filter_to="
    "&search=Search"
)

start_date = datetime(2025, 8, 11)
end_date = datetime(2025, 12, 31)

weekday_map = {
    0: "Monday",
    1: "Tuesday",
    2: "Wednesday",
    3: "Thursday",
    4: "Friday",
}

weekday_links = {
    "Monday": [],
    "Tuesday": [],
    "Wednesday": [],
    "Thursday": [],
    "Friday": [],
}

current = start_date
while current <= end_date:
    if current.weekday() in weekday_map:
        date_day = weekday_map[current.weekday()]
        date_str = current.strftime("%a, %d %b %Y")
        encoded_date = date_str.replace(",", "%2C").replace(" ", "+")
        weekday_links[date_day].append(
            base_url.format(date=encoded_date)
        )
    current += timedelta(days=1)

for index, day in weekday_map.items():
    print(f"{day}:")
    for link in weekday_links[day]:
        print(f"  {link}")
        print()
    print()  


