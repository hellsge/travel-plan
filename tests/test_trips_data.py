import trip_model as tm
from conftest import ROOT


def test_all_trip_files_valid():
    trips, errors = tm.load_trips(ROOT / "trips")
    assert errors == []
    assert trips
