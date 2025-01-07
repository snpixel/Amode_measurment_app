import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';

class UpdateMeasurementPage extends StatefulWidget {
  const UpdateMeasurementPage({Key? key}) : super(key: key);

  @override
  _UpdateMeasurementPageState createState() => _UpdateMeasurementPageState();
}

class _UpdateMeasurementPageState extends State<UpdateMeasurementPage> {
  final TextEditingController empNameController = TextEditingController();
  final TextEditingController sleeveController = TextEditingController();
  final TextEditingController chestController = TextEditingController();
  final TextEditingController shoulderController = TextEditingController();
  final TextEditingController lengthController = TextEditingController();

  String? selectedFile;
  List<Map<String, dynamic>> measurements = [];

  Future<List<String>> _getSavedFiles() async {
    Directory? directory = await getExternalStorageDirectory();
    String newPath = "";
    List<String> paths = directory!.path.split("/");
    for (int x = 1; x < paths.length; x++) {
      String folder = paths[x];
      if (folder != "Android") {
        newPath += "/$folder";
      } else {
        break;
      }
    }
    newPath = "$newPath/Download/amode_measurements";
    directory = Directory(newPath);

    if (await _requestPermissions()) {
      if (directory.existsSync()) {
        return directory
            .listSync()
            .where((file) => file.path.endsWith('.xlsx'))
            .map((file) => file.path.split('/').last)
            .toList();
      } else {
        return [];
      }
    } else {
      return [];
    }
  }

  Future<void> _appendToExistingFile(String fileName) async {
    Directory? directory = await getExternalStorageDirectory();
    String newPath = "";
    List<String> paths = directory!.path.split("/");
    for (int x = 1; x < paths.length; x++) {
      String folder = paths[x];
      if (folder != "Android") {
        newPath += "/$folder";
      } else {
        break;
      }
    }
    newPath = "$newPath/Download/amode_measurements";
    directory = Directory(newPath);

    if (await _requestPermissions()) {
      final filePath = "${directory.path}/$fileName";
      if (File(filePath).existsSync()) {
        // Read existing file
        final bytes = File(filePath).readAsBytesSync();
        var excel = Excel.decodeBytes(bytes);

        // Add data to Sheet1 if it exists, else create it
        var sheetName = "Sheet1";
        var sheet = excel.sheets[sheetName] ?? excel[sheetName];

        sheet.appendRow([
          empNameController.text.trim().isEmpty ? 'N/A' : empNameController.text.trim(),
          sleeveController.text,
          chestController.text,
          shoulderController.text,
          lengthController.text
        ]);

        // Write back
        File(filePath).writeAsBytesSync(excel.encode()!);

        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Data appended to $filePath')),
        );

        // Add the new measurement to the list
        setState(() {
          measurements.add({
            'Employee Name': empNameController.text.trim().isEmpty ? 'N/A' : empNameController.text.trim(),
            'Sleeve': sleeveController.text,
            'Chest': chestController.text,
            'Shoulder': shoulderController.text,
            'Length': lengthController.text
          });

          // Clear the input fields after adding measurement
          empNameController.clear();
          sleeveController.clear();
          chestController.clear();
          shoulderController.clear();
          lengthController.clear();
        });
      } else {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('File not found!')),
        );
      }
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Permission denied')),
      );
    }
  }

  Future<bool> _requestPermissions() async {
    if (await Permission.manageExternalStorage.request().isGranted) {
      return true;
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Permission denied')),
      );
      return false;
    }
  }

  void _showWarningDialog() {
    showDialog(
      context: context,
      builder: (BuildContext context) {
        return AlertDialog(
          backgroundColor: Colors.white,
          title: const Text('Confirm Measurement', style: TextStyle(color: Colors.black)),
          content: const Text('Are you sure you want to add this measurement?', style: TextStyle(color: Colors.black)),
          actions: <Widget>[
            TextButton(
              onPressed: () {
                Navigator.pop(context); // Close the dialog
              },
              child: const Text('Cancel', style: TextStyle(color: Colors.black)),
            ),
            TextButton(
              onPressed: () {
                if (selectedFile != null) {
                  _appendToExistingFile(selectedFile!);
                } else {
                  ScaffoldMessenger.of(context).showSnackBar(
                    const SnackBar(content: Text('Please select a file!')),
                  );
                }
                Navigator.pop(context); // Close the dialog
              },
              child: const Text('Yes', style: TextStyle(color: Colors.black)),
            ),
          ],
        );
      },
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Update Measurement File'),
        backgroundColor: Colors.white,
        foregroundColor: Colors.black,
      ),
      body: Center(
        child: SingleChildScrollView(
          child: Container(
            width: 300,
            padding: const EdgeInsets.all(16.0),
            child: Column(
              children: [
                FutureBuilder<List<String>>(
                  future: _getSavedFiles(),
                  builder: (context, snapshot) {
                    if (snapshot.connectionState == ConnectionState.waiting) {
                      return const CircularProgressIndicator();
                    } else if (snapshot.hasError) {
                      return const Text('Error loading files');
                    } else if (!snapshot.hasData || snapshot.data!.isEmpty) {
                      return const Text('No files found');
                    } else {
                      return Row(
                        children: [
                          Expanded(
                            child: DropdownButton<String>(
                              value: selectedFile,
                              hint: const Text('Select a file'),
                              items: snapshot.data!.map((String file) {
                                return DropdownMenuItem<String>(
                                  value: file,
                                  child: Text(file),
                                );
                              }).toList(),
                              onChanged: (String? newValue) {
                                setState(() {
                                  selectedFile = newValue;
                                });
                              },
                            ),
                          ),
                        ],
                      );
                    }
                  },
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: empNameController,
                  decoration: const InputDecoration(labelText: 'Employee Name'),
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: sleeveController,
                  keyboardType: TextInputType.number,
                  decoration: const InputDecoration(labelText: 'Sleeve (inches)'),
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: chestController,
                  keyboardType: TextInputType.number,
                  decoration: const InputDecoration(labelText: 'Chest (inches)'),
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: shoulderController,
                  keyboardType: TextInputType.number,
                  decoration: const InputDecoration(labelText: 'Shoulder (inches)'),
                ),
                const SizedBox(height: 10),
                TextField(
                  controller: lengthController,
                  keyboardType: TextInputType.number,
                  decoration: const InputDecoration(labelText: 'Length (inches)'),
                ),
                const SizedBox(height: 20),
                ElevatedButton(
                  onPressed: _showWarningDialog,
                  child: const Text('Add Measurement'),
                  style: ElevatedButton.styleFrom(backgroundColor: Colors.black),
                ),
                const SizedBox(height: 20),
                // Display measurements added below the button
                Column(
                  children: measurements.map((measurement) {
                    return Container(
                      margin: const EdgeInsets.symmetric(vertical: 10),
                      padding: const EdgeInsets.all(10),
                      decoration: BoxDecoration(
                        color: Colors.black, // Change box color to black
                        borderRadius: BorderRadius.circular(10),
                      ),
                      child: Text(
                        'Emp: ${measurement['Employee Name']}, Sleeve: ${measurement['Sleeve']}\", Chest: ${measurement['Chest']}\", Shoulder: ${measurement['Shoulder']}\", Length: ${measurement['Length']}\"',
                        style: const TextStyle(color: Colors.white), // Change text color to white
                      ),
                    );
                  }).toList(),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
