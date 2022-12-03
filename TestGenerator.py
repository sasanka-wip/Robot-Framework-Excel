

default_template = """
*** Settings ***
Library         Selenium2Library
Library         OperatingSystem
Library         BuiltIn
Library         Collections

Test Teardown     close browser

*** Test Cases ***
"""


def _write_to_file(data):
    with open("outfile.robot", "a") as outfile:
        outfile.write(data)


def _test_representation(TestName, TestData):
    _write_to_file("\n\n")
    _write_to_file(TestName)
    for test, test_steps in TestData.items():
        if test == TestName:
            for t in TestData[TestName]:
                raw_data = ('		'.join(map(str, t)).replace("None", ""))
                # print(raw_data)
                _write_to_file("\n")
                _write_to_file("\t%s" % raw_data)



