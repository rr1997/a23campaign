<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Get form data
    $name = $_POST['name'];
    $number = $_POST['number'];
    $email = $_POST['email'];
    $address = $_POST['address'];

    // Validate phone number and email
    if (!preg_match('/^\d{10}$/', $number)) {
        die('Invalid phone number. It must be 10 digits.');
    }
    if (!filter_var($email, FILTER_VALIDATE_EMAIL)) {
        die('Invalid email address.');
    }

    // Include PHPExcel library and create an instance
    require 'vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

    // Define the Excel file path
    $excelFilePath = 'user_data.xlsx';

    // Check if the Excel file exists
    if (file_exists($excelFilePath)) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($excelFilePath);
        $sheet = $spreadsheet->getActiveSheet();
    } else {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        // Create header
        $sheet->setCellValue('A1', 'Name');
        $sheet->setCellValue('B1', 'Phone Number');
        $sheet->setCellValue('C1', 'Email');
        $sheet->setCellValue('D1', 'Address');
    }

    // Find the next empty row
    $row = $sheet->getHighestRow() + 1;

    // Insert data
    $sheet->setCellValue('A' . $row, $name);
    $sheet->setCellValue('B' . $row, $number);
    $sheet->setCellValue('C' . $row, $email);
    $sheet->setCellValue('D' . $row, $address);

    // Save the Excel file
    $writer = new Xlsx($spreadsheet);
    $writer->save($excelFilePath);

    // HTML response
    echo "<html><head><title>Form Submission</title><style>
            body { font-family: Arial, sans-serif; background-color: #f0f0f0; margin: 0; padding: 20px; }
            .container { max-width: 500px; margin: auto; background: #fff; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); text-align: center; }
            h2 { color: #4CAF50; }
          </style></head><body><div class='container'><h2>Thank you! Your data has been submitted.</h2></div></body></html>";
}
?>
