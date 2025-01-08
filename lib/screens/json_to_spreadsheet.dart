import 'dart:convert';
import 'dart:io';
import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:path_provider/path_provider.dart';

class JsonToSpreadsheet extends StatefulWidget {
  @override
  _JsonToSpreadsheetState createState() => _JsonToSpreadsheetState();
}

class _JsonToSpreadsheetState extends State<JsonToSpreadsheet> {
  Future<String> getDownloadPath() async {
    Directory? directory = Directory('/storage/emulated/0/Download');
    if (!await directory.exists()) {
      directory = await getApplicationDocumentsDirectory();
    }
    return directory.path;
  }

  void pickJsonFile() async {
    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['json'],
      );

      if (result != null && result.files.single.path != null) {
        final file = File(result.files.single.path!);
        convertJsonToSpreadsheet(file);
      } else {
        ScaffoldMessenger.of(context).showSnackBar(SnackBar(
          content: Text("No file selected"),
        ));
      }
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        content: Text("Error selecting file: ${e.toString()}"),
      ));
    }
  }

  void convertJsonToSpreadsheet(File jsonFile) async {
    try {
      final jsonString = await jsonFile.readAsString();
      final jsonData = jsonDecode(jsonString) as List<dynamic>;

      if (jsonData.isEmpty) {
        throw Exception("JSON is empty");
      }

      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      final headers = jsonData.first.keys.toList();
      sheet.appendRow(headers.map((header) => header.toString()).toList());

      for (var item in jsonData) {
        final row = headers.map((header) {
          final value = item[header];
          if (value is int || value is double || value is String) {
            return value;
          }
          return value?.toString() ?? '';
        }).toList();
        sheet.appendRow(row);
      }

      final downloadPath = await getDownloadPath();
      final filePath = "$downloadPath/output.xlsx";
      final file = File(filePath);
      await file.writeAsBytes(excel.save()!);

      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        content: Text("Spreadsheet saved to $filePath"),
      ));
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        content: Text("Error: ${e.toString()}"),
      ));
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("JSON to Spreadsheet")),
      body: Center(
        child: ElevatedButton(
          onPressed: pickJsonFile,
          child: Text("Select JSON File"),
        ),
      ),
    );
  }
}
