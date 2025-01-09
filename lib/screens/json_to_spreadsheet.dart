import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:open_file/open_file.dart';
import 'package:permission_handler/permission_handler.dart';

class JsonToSpreadsheet extends StatefulWidget {
  const JsonToSpreadsheet({super.key});

  @override
  _JsonToSpreadsheetState createState() => _JsonToSpreadsheetState();
}

class _JsonToSpreadsheetState extends State<JsonToSpreadsheet> {
  /// Meminta izin penyimpanan pada perangkat Android.
  Future<bool> requestStoragePermission() async {
    final status = await Permission.manageExternalStorage.request();
    if (status.isGranted) {
      return true;
    } else {
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        content: Text("Permission denied. Please enable storage permission."),
      ));
      return false;
    }
  }

  /// Membuka File Picker untuk memilih file JSON.
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

  /// Mengonversi file JSON ke spreadsheet Excel.
  void convertJsonToSpreadsheet(File jsonFile) async {
    try {
      // Meminta izin penyimpanan
      if (!await requestStoragePermission()) return;

      // Membaca file JSON
      final jsonString = await jsonFile.readAsString();
      final jsonData = jsonDecode(jsonString) as List<dynamic>;

      if (jsonData.isEmpty) {
        throw Exception("JSON is empty");
      }

      // Membuat file Excel
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      // Menambahkan header ke Excel
      final headers = jsonData.first.keys.toList();
      sheet.appendRow(headers.map((header) => header.toString()).toList());

      // Menambahkan data ke Excel
      for (var item in jsonData) {
        final row = headers.map((header) {
          final value = item[header];
          return value?.toString() ?? '';
        }).toList();
        sheet.appendRow(row);
      }

      // Menyimpan file di direktori Downloads
      final bytes = Uint8List.fromList(excel.save()!);
      await saveToDownloads(bytes, "output.xlsx");
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        content: Text("Error: ${e.toString()}"),
      ));
    }
  }

  /// Menyimpan file di direktori Downloads.
  Future<void> saveToDownloads(Uint8List bytes, String fileName) async {
    try {
      final directory = Directory('/storage/emulated/0/Download');
      if (!await directory.exists()) {
        throw Exception("Downloads directory does not exist");
      }

      final file = File('${directory.path}/$fileName');
      await file.writeAsBytes(bytes);

      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        content: Text("Spreadsheet saved to ${file.path}"),
      ));

      // Membuka file setelah disimpan
      OpenFile.open(file.path);
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        content: Text("Error saving file: ${e.toString()}"),
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
