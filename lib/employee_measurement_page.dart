import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:intl/intl.dart';
import 'main.dart';

class EmployeeMeasurementPage extends StatefulWidget {
  final String companyName;

  const EmployeeMeasurementPage({required this.companyName});

  @override
  _EmployeeMeasurementPageState createState() => _EmployeeMeasurementPageState();
}

class _EmployeeMeasurementPageState extends State<EmployeeMeasurementPage> {
  final TextEditingController empNameController = TextEditingController();
  final TextEditingController mobileNumberController = TextEditingController();
  final TextEditingController departmentController = TextEditingController();
  final TextEditingController designationController = TextEditingController();
  final TextEditingController lengthController = TextEditingController();
  final TextEditingController shoulderController = TextEditingController();
  final TextEditingController sleeveController = TextEditingController();
  final TextEditingController chestController = TextEditingController();
  final TextEditingController waistController = TextEditingController();
  final TextEditingController bottomController = TextEditingController();
  final TextEditingController neckController = TextEditingController();
  final TextEditingController remarksController = TextEditingController();
  final TextEditingController pantWaistController = TextEditingController();
  final TextEditingController seatController = TextEditingController();
  final TextEditingController pantLengthController = TextEditingController();
  final TextEditingController innerLengthController = TextEditingController();
  final TextEditingController thighController = TextEditingController();
  final TextEditingController kneeController = TextEditingController();
  final TextEditingController pantBottomController = TextEditingController();
  final TextEditingController pantRemarksController = TextEditingController();

  String selectedMeasurementType = 'Shirt';
  List<Map<String, dynamic>> employeeData = [];

  Future<void> saveToExcel() async {
    if (await _requestPermissions()) {
      var excel = Excel.createExcel();
      Sheet sheetObject = excel['Sheet1'];
      // Add company name and creation date
      sheetObject.merge(CellIndex.indexByString("A1"), CellIndex.indexByString("B1"));
      sheetObject.cell(CellIndex.indexByString("A1")).value = widget.companyName.toUpperCase();
      sheetObject.cell(CellIndex.indexByString("C1")).value = DateFormat('yyyy-MM-dd HH:mm:ss').format(DateTime.now());

      // Merge cells for "Shirt" and "Pant"
      sheetObject.merge(CellIndex.indexByString("G1"), CellIndex.indexByString("N1"));
      var shirtCell = sheetObject.cell(CellIndex.indexByString("G1"));
      shirtCell.value = 'SHIRT';
      shirtCell.cellStyle = CellStyle(horizontalAlign: HorizontalAlign.Center,backgroundColorHex: "#BF8F00");

      sheetObject.merge(CellIndex.indexByString("P1"), CellIndex.indexByString("W1"));
var pantCell = sheetObject.cell(CellIndex.indexByString("P1"));
pantCell.value = 'PANT';
pantCell.cellStyle = CellStyle(horizontalAlign: HorizontalAlign.Center, backgroundColorHex: "#FF0000");

// Add headers
List<String> headers = [
  'SERIAL NUMBER', 'EMPLOYEE NAME', 'MOBILE NUMBER', 'DEPARTMENT', 'DESIGNATION','',
  'LENGTH', 'SHOULDER', 'SLEEVE', 'CHEST', 'WAIST', 'BOTTOM', 'NECK', 'REMARKS',
  '', 'WAIST', 'SEAT', 'LENGTH', 'INNER LENGTH', 'THIGH', 'KNEE', 'BOTTOM', 'REMARKS'
];

// Append headers
sheetObject.appendRow(headers);

// Apply styles to headers
for (int i = 0; i < headers.length; i++) {
  var cell = sheetObject.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 1));
  if (['LENGTH', 'SHOULDER', 'SLEEVE', 'CHEST', 'WAIST', 'BOTTOM', 'NECK', 'REMARKS'].contains(headers[i]) && i < 14) {
    cell.cellStyle = CellStyle(
      backgroundColorHex: "#8EA9DB", // Color for shirt columns
      horizontalAlign: HorizontalAlign.Center,
    );
  } else if (['WAIST', 'SEAT', 'LENGTH', 'INNER LENGTH', 'THIGH', 'KNEE', 'BOTTOM', 'REMARKS'].contains(headers[i]) && i > 13) {
    cell.cellStyle = CellStyle(
      backgroundColorHex: "#A9D08E", // Color for pant columns
      horizontalAlign: HorizontalAlign.Center,
    );
  } else {
    cell.cellStyle = CellStyle(
      backgroundColorHex: "#FFFFFF", // Default color for other columns
      horizontalAlign: HorizontalAlign.Center,
    );
  }
}

      // Add data
      for (int i = 0; i < employeeData.length; i++) {
        var data = employeeData[i];
        sheetObject.appendRow([
          i + 1,
          data['Employee Name'].toUpperCase(),
          data['Mobile Number'].toUpperCase(),
          data['Department'].toUpperCase(),
          data['Designation'].toUpperCase(),
          '',
          data['Length']?.toUpperCase() ?? 'N/A',
          data['Shoulder']?.toUpperCase() ?? 'N/A',
          data['Sleeve']?.toUpperCase() ?? 'N/A',
          data['Chest']?.toUpperCase() ?? 'N/A',
          data['Waist']?.toUpperCase() ?? 'N/A',
          data['Bottom']?.toUpperCase() ?? 'N/A',
          data['Neck']?.toUpperCase() ?? 'N/A',
          data['Shirt Remarks']?.toUpperCase() ?? 'N/A',
          '',
          data['Pant Waist']?.toUpperCase() ?? 'N/A',
          data['Seat']?.toUpperCase() ?? 'N/A',
          data['Pant Length']?.toUpperCase() ?? 'N/A',
          data['Inner Length']?.toUpperCase() ?? 'N/A',
          data['Thigh']?.toUpperCase() ?? 'N/A',
          data['Knee']?.toUpperCase() ?? 'N/A',
          data['Pant Bottom']?.toUpperCase() ?? 'N/A',
          data['Pant Remarks']?.toUpperCase() ?? 'N/A'
        ]);
      }

      // Get the directory to save the file
      Directory? directory = await getExternalStorageDirectory();
      String newPath = "";
      List<String> paths = directory!.path.split("/");
      for (int x = 1; x < paths.length; x++) {
        String folder = paths[x];
        if (folder != "Android") {
          newPath += "/" + folder;
        } else {
          break;
        }
      }
      newPath = newPath + "/Download/amode_measurements";
      directory = Directory(newPath);

      // Create the directory if it does not exist
      if (!directory.existsSync()) {
        directory.createSync(recursive: true);
      }

      // Save the file
      String filePath = "${directory.path}/${widget.companyName}_measurements.xlsx";
      File(filePath)
        ..createSync(recursive: true)
        ..writeAsBytesSync(excel.encode()!);

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Measurements saved to $filePath')),
      );
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Permission denied')),
      );
    }
  }

  Future<bool> _requestPermissions() async {
    if (await Permission.storage.request().isGranted) {
      return true;
    } else if (await Permission.manageExternalStorage.request().isGranted) {
      return true;
    } else {
      return false;
    }
  }

  void _showWarningDialog(String title, String content, VoidCallback onConfirm) {
    showDialog(
      context: context,
      builder: (BuildContext context) {
        return AlertDialog(
          backgroundColor: Colors.white,
          title: Text(title, style: TextStyle(color: Colors.black)),
          content: Text(content, style: TextStyle(color: Colors.black)),
          actions: <Widget>[
            TextButton(
              onPressed: () {
                Navigator.pop(context); // Close the dialog
              },
              child: const Text('Cancel', style: TextStyle(color: Colors.black)),
            ),
            TextButton(
              onPressed: () {
                onConfirm();
                Navigator.pop(context); // Close the dialog
              },
              child: const Text('Yes', style: TextStyle(color: Colors.black)),
            ),
          ],
        );
      },
    );
  }

  void _addData() {
    if (empNameController.text.trim().isEmpty ||
        departmentController.text.trim().isEmpty ||
        designationController.text.trim().isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Please fill all mandatory fields!')),
      );
      return;
    }

    if (selectedMeasurementType == 'Shirt') {
      if (lengthController.text.trim().isEmpty &&
          shoulderController.text.trim().isEmpty &&
          sleeveController.text.trim().isEmpty &&
          chestController.text.trim().isEmpty &&
          waistController.text.trim().isEmpty &&
          bottomController.text.trim().isEmpty &&
          neckController.text.trim().isEmpty) {
        lengthController.text = 'N/A';
        shoulderController.text = 'N/A';
        sleeveController.text = 'N/A';
        chestController.text = 'N/A';
        waistController.text = 'N/A';
        bottomController.text = 'N/A';
        neckController.text = 'N/A';
      } else if (lengthController.text.trim().isEmpty ||
          shoulderController.text.trim().isEmpty ||
          sleeveController.text.trim().isEmpty ||
          chestController.text.trim().isEmpty ||
          waistController.text.trim().isEmpty ||
          bottomController.text.trim().isEmpty ||
          neckController.text.trim().isEmpty) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Please fill all mandatory fields for Shirt!')),
        );
        return;
      }
    } else {
      if (pantWaistController.text.trim().isEmpty &&
          seatController.text.trim().isEmpty &&
          pantLengthController.text.trim().isEmpty &&
          innerLengthController.text.trim().isEmpty &&
          thighController.text.trim().isEmpty &&
          kneeController.text.trim().isEmpty &&
          pantBottomController.text.trim().isEmpty) {
        pantWaistController.text = 'N/A';
        seatController.text = 'N/A';
        pantLengthController.text = 'N/A';
        innerLengthController.text = 'N/A';
        thighController.text = 'N/A';
        kneeController.text = 'N/A';
        pantBottomController.text = 'N/A';
      } else if (pantWaistController.text.trim().isEmpty ||
          seatController.text.trim().isEmpty ||
          pantLengthController.text.trim().isEmpty ||
          innerLengthController.text.trim().isEmpty ||
          thighController.text.trim().isEmpty ||
          kneeController.text.trim().isEmpty ||
          pantBottomController.text.trim().isEmpty) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Please fill all mandatory fields for Pant!')),
        );
        return;
      }
    }

    _showWarningDialog(
      'Confirm Entry',
      'Do you want to save this entry?',
      () {
        setState(() {
          employeeData.add({
            'Employee Name': empNameController.text.trim().toUpperCase(),
            'Mobile Number': mobileNumberController.text.trim().toUpperCase(),
            'Department': departmentController.text.trim().toUpperCase(),
            'Designation': designationController.text.trim().toUpperCase(),
            'Length': lengthController.text.trim().isEmpty ? 'N/A' : lengthController.text.trim().toUpperCase(),
            'Shoulder': shoulderController.text.trim().isEmpty ? 'N/A' : shoulderController.text.trim().toUpperCase(),
            'Sleeve': sleeveController.text.trim().isEmpty ? 'N/A' : sleeveController.text.trim().toUpperCase(),
            'Chest': chestController.text.trim().isEmpty ? 'N/A' : chestController.text.trim().toUpperCase(),
            'Waist': waistController.text.trim().isEmpty ? 'N/A' : waistController.text.trim().toUpperCase(),
            'Bottom': bottomController.text.trim().isEmpty ? 'N/A' : bottomController.text.trim().toUpperCase(),
            'Neck': neckController.text.trim().isEmpty ? 'N/A' : neckController.text.trim().toUpperCase(),
            'Shirt Remarks': remarksController.text.trim().isEmpty ? 'N/A' : remarksController.text.trim().toUpperCase(),
            'Pant Waist': pantWaistController.text.trim().isEmpty ? 'N/A' : pantWaistController.text.trim().toUpperCase(),
            'Seat': seatController.text.trim().isEmpty ? 'N/A' : seatController.text.trim().toUpperCase(),
            'Pant Length': pantLengthController.text.trim().isEmpty ? 'N/A' : pantLengthController.text.trim().toUpperCase(),
            'Inner Length': innerLengthController.text.trim().isEmpty ? 'N/A' : innerLengthController.text.trim().toUpperCase(),
            'Thigh': thighController.text.trim().isEmpty ? 'N/A' : thighController.text.trim().toUpperCase(),
            'Knee': kneeController.text.trim().isEmpty ? 'N/A' : kneeController.text.trim().toUpperCase(),
            'Pant Bottom': pantBottomController.text.trim().isEmpty ? 'N/A' : pantBottomController.text.trim().toUpperCase(),
            'Pant Remarks': pantRemarksController.text.trim().isEmpty ? 'N/A' : pantRemarksController.text.trim().toUpperCase(),
          });

          // Clear the input fields after adding data
          empNameController.clear();
          mobileNumberController.clear();
          departmentController.clear();
          designationController.clear();
          lengthController.clear();
          shoulderController.clear();
          sleeveController.clear();
          chestController.clear();
          waistController.clear();
          bottomController.clear();
          neckController.clear();
          remarksController.clear();
          pantWaistController.clear();
          seatController.clear();
          pantLengthController.clear();
          innerLengthController.clear();
          thighController.clear();
          kneeController.clear();
          pantBottomController.clear();
          pantRemarksController.clear();
        });
      },
    );
  }

  void _endSession() {
    _showWarningDialog(
      'End Session',
      'Are you sure you want to end the session?',
      () async {
        await saveToExcel();
        setState(() {
          employeeData.clear();
        });
        Navigator.pushReplacement(
          context,
          MaterialPageRoute(builder: (context) => const HomeScreen()), // Go back to HomePage
        );
      },
    );
  }
  
  @override
  Widget build(BuildContext context) {
    return WillPopScope(
    onWillPop: _onWillPop,
    child: Scaffold(
      appBar: AppBar(
        title: const Text('Employee and Measurement Data Entry'),
        automaticallyImplyLeading: false,
        backgroundColor: Colors.white,
        foregroundColor: Colors.black,
      ),
      body: Center(
        child: SingleChildScrollView(
          child: Container(
            width: 450,
            padding: const EdgeInsets.all(16.0),
            child: Column(
              children: [
                // Employee Data Input Box
                TextField(
                  controller: empNameController,
                  decoration: const InputDecoration(labelText: 'Employee Name (mandatory)'),
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: mobileNumberController,
                  keyboardType: TextInputType.phone,
                  decoration: const InputDecoration(labelText: 'Mobile Number (optional)'),
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: departmentController,
                  decoration: const InputDecoration(labelText: 'Department (mandatory)'),
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: designationController,
                  decoration: const InputDecoration(labelText: 'Designation (mandatory)'),
                ),
                const SizedBox(height: 20),
                // Measurement Input Box
                DropdownButton<String>(
                  value: selectedMeasurementType,
                  items: <String>['Shirt', 'Pant'].map((String value) {
                    return DropdownMenuItem<String>(
                      value: value,
                      child: Text(value),
                    );
                  }).toList(),
                  onChanged: (String? newValue) {
                    setState(() {
                      selectedMeasurementType = newValue!;
                    });
                  },
                ),
                const SizedBox(height: 10),
                if (selectedMeasurementType == 'Shirt') ...[
                  Row(
                    children: [
                      Expanded(
                        child: Column(
                          children: [
                            TextField(
                              controller: lengthController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Length'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: sleeveController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Sleeve'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: waistController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Waist'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: neckController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Neck'),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(width: 20),
                      Expanded(
                        child: Column(
                          children: [
                            TextField(
                              controller: shoulderController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Shoulder'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: chestController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Chest'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: bottomController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Bottom'),
                            ),
                            const SizedBox(height: 62), // Add extra space to align with the left column
                          ],
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 10),
                  TextField(
                    controller: remarksController,
                    decoration: const InputDecoration(labelText: 'Remarks (optional)'),
                  ),
                ] else ...[
                  Row(
                    children: [
                      Expanded(
                        child: Column(
                          children: [
                            TextField(
                              controller: pantWaistController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Waist'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: pantLengthController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Length'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: thighController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Thigh'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: pantBottomController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Bottom'),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(width: 20),
                      Expanded(
                        child: Column(
                          children: [
                            TextField(
                              controller: seatController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Seat'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: innerLengthController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Inner Length'),
                            ),
                            const SizedBox(height: 10),
                            TextField(
                              controller: kneeController,
                              keyboardType: TextInputType.number,
                              decoration: const InputDecoration(labelText: 'Knee'),
                            ),
                            const SizedBox(height: 61), // Add extra space to align with the left column
                          ],
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 10),
                  TextField(
                    controller: pantRemarksController,
                    decoration: const InputDecoration(labelText: 'Remarks (optional)'),
                  ),
                ],
                const SizedBox(height: 20),
                ElevatedButton(
                  onPressed: _addData,
                  child: const Text('Add Data'),
                  style: ElevatedButton.styleFrom(backgroundColor: Colors.black),
                ),
                const SizedBox(height: 20),
                ElevatedButton(
                  onPressed: _endSession,
                  child: const Text('End Session'),
                  style: ElevatedButton.styleFrom(backgroundColor: Colors.black),
                ),
                const SizedBox(height: 20),
                // Display entered data
                if (employeeData.isNotEmpty)
                  Container(
                    height: 200,
                    width: double.infinity,
                    child: ListView.builder(
                      itemCount: employeeData.length,
                      itemBuilder: (context, index) {
                        var data = employeeData[index];
                        String measurementType = '';
                        if (data['Length'] != 'N/A' || data['Shoulder'] != 'N/A' || data['Sleeve'] != 'N/A') {
                          measurementType = 'Shirt';
                        }
                        if (data['Pant Waist'] != 'N/A' || data['Seat'] != 'N/A' || data['Pant Length'] != 'N/A') {
                          if (measurementType.isNotEmpty) {
                            measurementType += ' & Pant';
                          } else {
                            measurementType = 'Pant';
                          }
                        }
                        return Container(
                          margin: const EdgeInsets.symmetric(vertical: 5),
                          padding: const EdgeInsets.all(10),
                          decoration: BoxDecoration(
                            color: Colors.grey[200],
                            borderRadius: BorderRadius.circular(10),
                          ),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text('Employee Name: ${data['Employee Name']}'),
                              Text('Designation: ${data['Designation']}'),
                              Text('Department: ${data['Department']}'),
                              Text('Measurement Type: $measurementType'),
                            ],
                          ),
                        );
                      },
                    ),
                  ),
              ],
            ),
          ),
        ),
      ),
    ));
  }
  Future<bool> _onWillPop() async {
  if (employeeData.isEmpty) {
    return true; // Allow back navigation if no data entered
  }

  return await showDialog<bool>(
    context: context,
    builder: (BuildContext context) {
      return AlertDialog(
        backgroundColor: Colors.white,
        title: Text('Warning', style: TextStyle(color: Colors.black)),
        content: Text(
          'You have unsaved data. Are you sure you want to go back? This will discard all entered data.',
          style: TextStyle(color: Colors.black)
        ),
        actions: <Widget>[
          TextButton(
            onPressed: () {
              Navigator.of(context).pop(false); // Stay on page
            },
            child: Text('Cancel', style: TextStyle(color: Colors.black)),
          ),
          TextButton(
            onPressed: () async {
              await saveToExcel(); // Save data before leaving
              Navigator.of(context).pop(true); // Allow back navigation
              Navigator.pushReplacement(
                context,
                MaterialPageRoute(builder: (context) => const HomeScreen()),
              );
            },
            child: Text('Yes', style: TextStyle(color: Colors.black)),
          ),
        ],
      );
    },
  ) ?? false; // If dialog is dismissed by tapping outside, default to false
}
}
