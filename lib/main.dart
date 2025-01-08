import 'package:flutter/material.dart';
import 'screens/json_to_spreadsheet.dart';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'JSON to Spreadsheet',
      theme: ThemeData(primarySwatch: Colors.blue),
      home: JsonToSpreadsheet(),
    );
  }
}
