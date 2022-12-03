import os
import time
from datetime import datetime
from subprocess import call
from Reader import TestReader
from TestGenerator import _test_representation, _write_to_file, default_template
from robot.api import ExecutionResult, ResultVisitor
from Resultwriter import ResultWriter

##################################################
timestamp = time.strftime('%Ihr-%Mmin')
current_month_text = datetime.now().strftime('%h')
current_year = datetime.now().strftime('%Y')
current_day = datetime.now().strftime('%d')
##################################################
result_dict = {}

class TestMetrics(ResultVisitor):

    def visit_test(self, test):
        result_dict.update({test.name: test.status})

if __name__ == "__main__":
    test_excel_file = 'Automation_Runner.xlsx'
    # Delete outfile.robot before test
    robot_testfile = "outfile.robot"
    try:
        os.remove(robot_testfile)
    except OSError as e:
        print("Error: %s - %s." % (e.filename, e.strerror))
    test_data = TestReader(test_excel_file).testcase_formation()
    testcases_name = [key for key in test_data.keys()]
    tests = [x for x in test_data if x is not None]
    _write_to_file(default_template)
    for test_name in testcases_name:
        _test_representation(test_name, test_data)

    cur_dir = os.path.dirname(os.path.realpath(__file__))
    result_dir = cur_dir + "\\" + "output" + "\\" + current_year + "\\" + current_month_text + "-" + current_day + "\\" + timestamp

    print("#############################")
    print("Running Test.")
    print("#############################")

    call(["robot", "--outputdir", result_dir, robot_testfile])

    #NOTE: As we are fetching result from output.xml after execution, And we are not sure when the
    #output.xml file will be availabe. so adding below condition to wait till file availabe.

    time_to_wait = 43200  #Max 12 hr
    time_counter = 0
    file_path = result_dir + "\\" + "output.xml"
    while not os.path.exists(file_path):
        time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait: break

    result = ExecutionResult(result_dir + "\\" + "output.xml")
    result.configure(stat_config={'suite_stat_level': 2,
                              'tag_stat_combine': 'tagANDanother'})
    result.visit(TestMetrics())

    for i in result_dict.items():
        cords = ResultWriter(test_excel_file).get_coordinates(i[0])
        report = result_dir + "\\" + "report.html"
        ResultWriter(test_excel_file).add_value(cords, i[1],report)

