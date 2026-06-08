def pytest_runtest_logreport(report):
    if report.when == "call":
        print()
