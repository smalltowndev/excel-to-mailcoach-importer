<?php
require 'vendor/autoload.php';

use Spatie\MailcoachSdk\Mailcoach;
use PhpOffice\PhpSpreadsheet\IOFactory;

error_reporting(E_ERROR);

$dotenv = Dotenv\Dotenv::createImmutable(__DIR__);
$dotenv->safeLoad();

$mailcoach_api_key = $_ENV['MAILCOACH_API_KEY'];
$mailcoach_base_url = $_ENV['MAILCOACH_BASE_URL'];

function readExcelFile($filePath) {
    try {
        // Load the spreadsheet file
        $spreadsheet = IOFactory::load($filePath);

        // Get the active sheet
        $sheet = $spreadsheet->getActiveSheet();

        // Get the highest row and column
        $highestRow = $sheet->getHighestRow();
        $highestColumn = $sheet->getHighestColumn();

        $data = [];

        // Loop through each row and column
        for ($row = 1; $row <= $highestRow; $row++) {
            $rowData = [];
            for ($col = 'A'; $col <= $highestColumn; $col++) {
                $rowData[] = $sheet->getCell($col . $row)->getValue();
            }
            $data[] = $rowData;
        }

        return $data;
    } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
        echo 'Error loading file: ' . $e->getMessage();
        return [];
    }
}

/**
 * Replace with your XLSX file path
 *
 * We are only reading first three columns and they should be
 * email, first_name and last_name.
 */
$filePath = 'data/file-to-import.xlsx';
$data = readExcelFile($filePath);
$list_key = 'list-2';

// Output the data
foreach ( $data as $row ) {
    if ( 'email' === $row[0] ) {
        continue;
    }

    $email = $row[0];
    $email = filter_var($email, FILTER_VALIDATE_EMAIL);
    if ( ! $email ) {
        continue;
    }
    $first_name = $row[1];
    $last_name = $row[2];

    $mailcoach = new Mailcoach($mailcoach_api_key, $mailcoach_base_url );

    // Your Mailcoach list ids
    $mailcoach_list = array(
        'list-1' => '8903d129-64e8-4e4c-a32e-0d383df764fd',
        'list-2' => '3da2d7a5-6b3b-46b4-ab09-cc5efcacs2d0',
    );

    if ( isset( $mailcoach_list[ $list_key ] ) ) {
        $existing_subscriber = $mailcoach
            ->emailList( $mailcoach_list[ $list_key ] )
            ->subscriber( $email );
        $is_subscribed_already = (bool) $existing_subscriber;

        if ( ! $is_subscribed_already ) {
            // Add to List.
            $subscriber = $mailcoach->createSubscriber( $mailcoach_list[ $list_key ], [
                'email'     => $email,
                'first_name' => $first_name,
                'last_name' => $last_name,
            ]);

            echo 'Subscribed email: ' . $email . PHP_EOL;
        }

        // Run only when subscribing to 2nd list.
        // Which then removes the user from 1st list.
        if ( 'list-2' === $list_key ) {
            // Check if user is subscribed to first list.
            $existing_subscriber = $mailcoach
                ->emailList( $mailcoach_list[ 'list-1' ] )
                ->subscriber( $email );

            $is_subscribed_to_list = (bool) $existing_subscriber;

            if ( $is_subscribed_to_list ) {
                // Unsubscribe from first list.
                $existing_subscriber->delete();
                echo 'Deleted from free: ' . $email . PHP_EOL;
            }
        }

        // Sleep for 4 seconds to avoid throwing too many requests at Mailcoach while importing.
        sleep(4);
    }
}