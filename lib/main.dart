import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:path/path.dart' as p;
import 'dart:io';
import 'dart:typed_data';
import 'package:share_plus/share_plus.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    WidgetsFlutterBinding.ensureInitialized();
    return MaterialApp(
      home: Scaffold(
        appBar: AppBar(
          title: const Text('Share Excel Example'),
        ),
        body: Center(
          child: ElevatedButton(
            onPressed: () async {
              await createExcel(context);
            },
            child: const Text('Share Excel File'),
          ),
        ),
      ),
    );
  }
}

Future<void> createExcel(context) async {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];
  String persenate = '0';
  // Membuat header untuk 70 kolom
  for (int col = 0; col < 70; col++) {
    sheet
        .cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: col))
        .value = "Kolom $col";
  }
  // Membuat 6000 baris data contoh
  for (int row = 1; row <= 6000; row++) {
    print('loading: ${row / 6000 / 100}%');
    for (int col = 0; col < 70; col++) {
      sheet
          .cell(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: col))
          .value = "Data $row-$col";
    }
  }

  // Menyimpan file Excel ke perangkat
  var directory = await getApplicationDocumentsDirectory();
  var filePath = p.join(directory.path, 'data.xlsx');
  File(filePath)
    ..createSync(recursive: true)
    ..writeAsBytesSync(excel.encode()!);

  print("File Excel berhasil dibuat di $filePath");

  // Berbagi file Excel
  await Share.shareXFiles(
    [XFile(filePath)],
    text: 'Share file: $filePath',
  );
}
