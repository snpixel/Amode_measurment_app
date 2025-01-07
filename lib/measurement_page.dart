import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'main.dart';

class MeasurementPage extends StatefulWidget {
  final String companyName;

  const MeasurementPage({required this.companyName});

  @override
  _MeasurementPageState createState() => _MeasurementPageState();
}

class _MeasurementPageState extends State<MeasurementPage> {
  final TextEditingController empNameController = TextEditingController();
  final TextEditingController sleeveController = TextEditingController();
  final TextEditingController chestController = TextEditingController();
  final TextEditingController shoulderController = TextEditingController();
  final TextEditingController lengthController = TextEditingController();

  List<Map<String, dynamic>> measurements = [];

  Future<void> saveToExcel() async {
    // Request storage permissions
    if (await _requestPermissions()) {
      var excel = Excel.createExcel();
      Sheet sheetObject = excel['Sheet1'];

      // Add company name as the first row
      sheetObject.appendRow([widget.companyName]);

      // Add headers
      sheetObject.appendRow(['Employee Name', 'Sleeve', 'Chest', 'Shoulder', 'Length']);

      // Add data
      for (var measurement in measurements) {
        sheetObject.appendRow([
          measurement['Employee Name'],
          measurement['Sleeve'],
          measurement['Chest'],
          measurement['Shoulder'],
          measurement['Length']
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

  Future<bool> _onWillPop() async {
    return (await showDialog(
      context: context,
      builder: (context) => AlertDialog(
        backgroundColor: Colors.white,
        title: Text('Warning', style: TextStyle(color: Colors.black)),
        content: Text('Are you sure you want to go back? All data will be lost.', style: TextStyle(color: Colors.black)),
        actions: <Widget>[
          TextButton(
            onPressed: () => Navigator.of(context).pop(false),
            child: Text('Cancel', style: TextStyle(color: Colors.black)),
          ),
          TextButton(
            onPressed: () {
              setState(() {
                measurements.clear();
              });
              Navigator.of(context).pop(true);
            },
            child: Text('Yes', style: TextStyle(color: Colors.black)),
          ),
        ],
      ),
    )) ?? false;
  }

  @override
  Widget build(BuildContext context) {
    return WillPopScope(
      onWillPop: _onWillPop,
      child: Scaffold(
        appBar: AppBar(
          title: const Text('Measurements Page'),
          automaticallyImplyLeading: false, // Remove the back button
          backgroundColor: Colors.white,
          foregroundColor: Colors.black,
        ),
        body: LayoutBuilder(
          builder: (context, constraints) {
            return Center(
              child: SingleChildScrollView(
                child: Padding(
                  padding: const EdgeInsets.symmetric(vertical: 20.0),
                  child: Container(
                    width: constraints.maxWidth * 0.85,  // Use 85% of the screen width
                    padding: const EdgeInsets.all(20.0),
                    decoration: BoxDecoration(
                      color: Colors.white,
                      borderRadius: BorderRadius.circular(10),
                    ),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.center,
                      children: [
                        // Input Fields for Measurements
                        TextField(
                          controller: empNameController,
                          style: TextStyle(color: Colors.black),
                          decoration: const InputDecoration(
                            labelText: 'Employee Name (optional)',
                            labelStyle: TextStyle(color: Colors.black),
                            border: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            enabledBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            focusedBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                          ),
                        ),
                        const SizedBox(height: 10),
                        TextField(
                          controller: sleeveController,
                          keyboardType: TextInputType.number,
                          style: TextStyle(color: Colors.black),
                          decoration: const InputDecoration(
                            labelText: 'Sleeve (inches)',
                            labelStyle: TextStyle(color: Colors.black),
                            border: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            enabledBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            focusedBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                          ),
                        ),
                        const SizedBox(height: 10),
                        TextField(
                          controller: chestController,
                          keyboardType: TextInputType.number,
                          style: TextStyle(color: Colors.black),
                          decoration: const InputDecoration(
                            labelText: 'Chest (inches)',
                            labelStyle: TextStyle(color: Colors.black),
                            border: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            enabledBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            focusedBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                          ),
                        ),
                        const SizedBox(height: 10),
                        TextField(
                          controller: shoulderController,
                          keyboardType: TextInputType.number,
                          style: TextStyle(color: Colors.black),
                          decoration: const InputDecoration(
                            labelText: 'Shoulder (inches)',
                            labelStyle: TextStyle(color: Colors.black),
                            border: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            enabledBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            focusedBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                          ),
                        ),
                        const SizedBox(height: 10),
                        TextField(
                          controller: lengthController,
                          keyboardType: TextInputType.number,
                          style: TextStyle(color: Colors.black),
                          decoration: const InputDecoration(
                            labelText: 'Length (inches)',
                            labelStyle: TextStyle(color: Colors.black),
                            border: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            enabledBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                            focusedBorder: OutlineInputBorder(
                              borderSide: BorderSide(color: Colors.black),
                            ),
                          ),
                        ),
                        const SizedBox(height: 20),
                        // Add Measurement Button
                        ElevatedButton(
                          onPressed: () {
                            String empName = empNameController.text.trim().isEmpty
                                ? 'N/A'
                                : empNameController.text.trim();
                            double? sleeve = double.tryParse(sleeveController.text);
                            double? chest = double.tryParse(chestController.text);
                            double? shoulder = double.tryParse(shoulderController.text);
                            double? length = double.tryParse(lengthController.text);

                            if (sleeve != null && chest != null && shoulder != null && length != null) {
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
                                          setState(() {
                                            measurements.add({
                                              'Employee Name': empName,
                                              'Sleeve': sleeve,
                                              'Chest': chest,
                                              'Shoulder': shoulder,
                                              'Length': length,
                                            });
                                            // Clear the input fields after adding measurement
                                            empNameController.clear();
                                            sleeveController.clear();
                                            chestController.clear();
                                            shoulderController.clear();
                                            lengthController.clear();
                                          });
                                          Navigator.pop(context); // Close the dialog
                                        },
                                        child: const Text('Yes', style: TextStyle(color: Colors.black)),
                                      ),
                                    ],
                                  );
                                },
                              );
                            } else {
                              ScaffoldMessenger.of(context).showSnackBar(
                                const SnackBar(content: Text('Please enter valid measurements!')),
                              );
                            }
                          },
                          child: const Text('Add Measurement', style: TextStyle(color: Colors.white)),
                          style: ElevatedButton.styleFrom(
                            backgroundColor: Colors.black,
                          ),
                        ),
                        const SizedBox(height: 20),
                        // End Session Button
                        ElevatedButton(
                          onPressed: () {
                            // Show warning dialog before ending session
                            showDialog(
                              context: context,
                              builder: (BuildContext context) {
                                return AlertDialog(
                                  backgroundColor: Colors.white,
                                  title: const Text('End Session', style: TextStyle(color: Colors.black)),
                                  content: const Text('Are you sure you want to end the session?', style: TextStyle(color: Colors.black)),
                                  actions: <Widget>[
                                    TextButton(
                                      onPressed: () {
                                        Navigator.pop(context); // Close the dialog
                                      },
                                      child: const Text('Cancel', style: TextStyle(color: Colors.black)),
                                    ),
                                    TextButton(
                                      onPressed: () async {
                                        await saveToExcel();
                                        setState(() {
                                          measurements.clear();
                                        });
                                        Navigator.pop(context); // Close the dialog
                                        Navigator.pushReplacement(
                                          context,
                                          MaterialPageRoute(builder: (context) => const HomeScreen()), // Go back to HomePage
                                        );
                                      },
                                      child: const Text('Yes', style: TextStyle(color: Colors.black)),
                                    ),
                                  ],
                                );
                              },
                            );
                          },
                          child: const Text('End Session', style: TextStyle(color: Colors.white)),
                          style: ElevatedButton.styleFrom(
                            backgroundColor: Colors.black,
                          ),
                        ),
                        const SizedBox(height: 20),
                        // Display measurements added below the buttons
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
          },
        ),
      ),
    );
  }
}
