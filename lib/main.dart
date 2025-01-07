import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:open_file/open_file.dart';
import 'dart:io';
import 'company_name_page.dart';
import 'update_employee_measurement_page.dart';

void main() {
  runApp(const MyApp()); // Add const
}

class MyApp extends StatelessWidget {
  const MyApp(); // Add const

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Session App',
      theme: ThemeData.light().copyWith(
        primaryColor: Colors.white,
        scaffoldBackgroundColor: Colors.white,
        appBarTheme: const AppBarTheme(
          backgroundColor: Colors.white,
          foregroundColor: Colors.black,
        ),
        elevatedButtonTheme: ElevatedButtonThemeData(
          style: ElevatedButton.styleFrom(
            backgroundColor: Colors.black,
            foregroundColor: Colors.white,
          ),
        ),
        inputDecorationTheme: const InputDecorationTheme(
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
      home: const HomeScreen(),
    );
  }
}

// Home Page
class HomeScreen extends StatelessWidget {
  const HomeScreen(); // Add const

  Future<List<FileSystemEntity>> _getRecentFiles() async {
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

    if (await Permission.manageExternalStorage.request().isGranted) {
      if (directory.existsSync()) {
        return directory.listSync().where((file) => file.path.endsWith('.xlsx')).toList();
      } else {
        return [];
      }
    } else {
      return [];
    }
  }

  void _openFile(BuildContext context, File file) async {
    final result = await OpenFile.open(file.path);
    if (result.type != ResultType.done) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Failed to open ${file.path}')),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      // Remove the AppBar
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            ElevatedButton(
              onPressed: () {
                Navigator.push(
                  context,
                  MaterialPageRoute(builder: (context) => CompanyNamePage()),
                );
              },
              child: const Text(
                'Start Session',
                style: TextStyle(
                  color: Colors.white, // Change text color to white
                  fontSize: 18, // Increase text size
                ),
              ),
              style: ElevatedButton.styleFrom(
                backgroundColor: Colors.black,
                padding: const EdgeInsets.symmetric(horizontal: 50, vertical: 20), // Increase button size
              ),
            ),
            const SizedBox(height: 20),
            ElevatedButton(
              onPressed: () {
                Navigator.push(
                  context,
                  MaterialPageRoute(builder: (context) => UpdateEmployeeMeasurementPage()),
                );
              },
              child: const Text(
                'Update File',
                style: TextStyle(
                  color: Colors.white, // Change text color to white
                  fontSize: 18, // Increase text size
                ),
              ),
              style: ElevatedButton.styleFrom(
                backgroundColor: Colors.black,
                padding: const EdgeInsets.symmetric(horizontal: 50, vertical: 20), // Increase button size
              ),
            ),
            const SizedBox(height: 20),
            IconButton(
              icon: const Icon(Icons.history, color: Colors.black),
              onPressed: () async {
                List<FileSystemEntity> recentFiles = await _getRecentFiles();
                showDialog(
                  context: context,
                  builder: (BuildContext context) {
                    return AlertDialog(
                      backgroundColor: Colors.white,
                      title: Text('Recent Files', style: TextStyle(color: Colors.black)),
                      content: Container(
                        width: double.maxFinite,
                        child: ListView.builder(
                          shrinkWrap: true,
                          itemCount: recentFiles.length,
                          itemBuilder: (context, index) {
                            File file = recentFiles[index] as File;
                            return ListTile(
                              title: Text(file.path.split('/').last, style: TextStyle(color: Colors.black)),
                              onTap: () {
                                Navigator.pop(context);
                                _openFile(context, file);
                              },
                            );
                          },
                        ),
                      ),
                      actions: [
                        TextButton(
                          onPressed: () {
                            Navigator.pop(context);
                          },
                          child: Text('Close', style: TextStyle(color: Colors.black)),
                        ),
                      ],
                    );
                  },
                );
              },
            ),
          ],
        ),
      ),
    );
  }
}
