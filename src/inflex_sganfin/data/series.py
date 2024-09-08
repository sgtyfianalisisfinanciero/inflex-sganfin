# class Unit:
#     def __init__(self) -> None:
#         symbol: str = ""


# class SeriesLocation:
#     def __init__(self) -> None:
#         timestamp: str = ""
#         offset: str = ""
#         offset_type: str = ""  # absolute or relative
#         value_type: str = ""  # value, rolling_avg,...


class Series:
    def __init__(self, report_name: str) -> None:
        self.report_name: str = report_name
        # agnostic_id: str = ""
        # unit: Unit = None
