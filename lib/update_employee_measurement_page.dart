import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'dart:io';

class UpdateEmployeeMeasurementPage extends StatefulWidget {
  @override
  _UpdateEmployeeMeasurementPageState createState() => _UpdateEmployeeMeasurementPageState();
}

class _UpdateEmployeeMeasurementPageState extends State<UpdateEmployeeMeasurementPage> {
  String? selectedFile;
  List<String> files = [];
  bool fileSelected = false;

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

  @override
  void initState() {
    super.initState();
    _listFiles();
  }

  Future<void> _listFiles() async {
    if (await Permission.manageExternalStorage.request().isGranted) {
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

      if (directory.existsSync()) {
        List<FileSystemEntity> fileEntities = await directory.list().toList();
        setState(() {
          files = fileEntities
              .where((file) => file.path.endsWith('.xlsx'))
              .map((file) => file.path.split('/').last)
              .toList();
        });
      }
    }
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
  }

  Future<void> _endSession() async {
    if (selectedFile == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('No file selected')),
      );
      return;
    }

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

    String filePath = "${directory.path}/$selectedFile";
    var bytes = File(filePath).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    Sheet sheetObject = excel['Sheet1'];

    // Add data
    for (int i = 0; i < employeeData.length; i++) {
      var data = employeeData[i];
      sheetObject.appendRow([
        i + 1,
        data['Employee Name'].toUpperCase(),
        data['Mobile Number'].toUpperCase(),
        data['Department'].toUpperCase(),
        data['Designation'].toUpperCase(),
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

    File(filePath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(excel.encode()!);

    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text('Measurements saved to $filePath')),
    );

    setState(() {
      employeeData.clear();
    });

    Navigator.pop(context);
  }

  @override
  Widget build(BuildContext context) {
    return WillPopScope(
    onWillPop: () => _onWillPop(context),
    child: Scaffold(
      appBar: AppBar(
        title: const Text('Update Employee Measurement Data'),
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
                DropdownButton<String>(
                  value: selectedFile,
                  hint: const Text('Select File'),
                  items: files.map((String file) {
                    return DropdownMenuItem<String>(
                      value: file,
                      child: Text(file),
                    );
                  }).toList(),
                  onChanged: (String? newValue) {
                    setState(() {
                      selectedFile = newValue;
                      fileSelected = true;
                    });
                  },
                ),
                if (selectedFile != null) ...[
                  Text('Selected File: $selectedFile'),
                  const SizedBox(height: 20),
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
                              borderRadius: BorderRadius.circular(5),
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
              ],
            ),
          ),
        ),
      ),
    ));
  }
  Future<bool> _onWillPop(BuildContext context) async {
  if (employeeData.isEmpty) {
    return true; // Allow back if no data entered
  }

  final shouldPop = await showDialog<bool>(
    context: context,
    builder: (context) {
      return AlertDialog(
        title: const Text('Are you sure?'),
        content: const Text('You have unsaved data. Going back will lose all entered measurements.'),
        actions: [
          TextButton(
            onPressed: () {
              Navigator.of(context).pop(false); // Stay on page
            },
            child: const Text('Cancel'),
          ),
          TextButton(
            onPressed: () {
              Navigator.of(context).pop(true); // Leave page
            },
            child: const Text('Leave anyway'),
          ),
        ],
      );
    },
  );

  return shouldPop ?? false;
}
}
