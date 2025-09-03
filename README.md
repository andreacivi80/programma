<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Programma di Produzione</title>
    <link href="https://fonts.googleapis.com/css2?family=Quicksand:wght@400;500;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Quicksand:wght@400;500;700&display=swap');

        body {
            font-family: 'Quicksand', sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #e0f7fa;
            color: #333;
            display: flex;
            justify-content: center;
            align-items: flex-start;
            min-height: 100vh;
            box-sizing: border-box;
        }

        .login-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.6);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 2000;
        }

        .login-container {
            background-color: #fff;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.2);
            text-align: center;
            max-width: 400px;
            width: 90%;
            display: flex;
            flex-direction: column;
            gap: 20px;
        }

        .login-container h2 {
            margin: 0;
            color: #3F51B5;
            font-size: 1.8em;
        }

        .login-container p {
            margin: 0;
            color: #555;
        }

        .login-container input {
            padding: 12px;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 1em;
            text-align: center;
        }

        .login-container button {
            padding: 12px 25px;
            border: none;
            border-radius: 8px;
            background-color: #4CAF50;
            color: white;
            font-size: 1em;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .login-container button:hover {
            background-color: #43a047;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }

        .login-container #loginError {
            color: red;
            display: none;
            font-size: 0.9em;
        }

        .container {
            background-color: #ffffff;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);
            padding: 30px;
            width: 98%;
            max-width: 1900px;
            box-sizing: border-box;
            overflow-x: hidden;
        }

        header {
            display: block;
            border-bottom: none;
            padding-bottom: 0;
            margin-bottom: 25px;
        }

        .header-layout-table {
            width: 100%;
            border-collapse: collapse;
            border: 1px solid #c8e6c9;
        }

        .header-layout-table td {
            border: 1px solid #c8e6c9;
            padding: 10px;
            vertical-align: top;
        }

        .header-logo-cell {
            width: 15%;
            text-align: center;
            padding: 5px;
        }
        .header-logo-cell img {
            max-width: 100px;
            height: auto;
            display: block;
            margin: 0 auto;
        }

        .header-top-center-cell,
        .header-bottom-center-cell {
            width: 45%;
            text-align: center;
        }

        .header-top-right-cell,
        .header-bottom-right-cell {
            width: 40%;
            text-align: right;
        }

        .header-text-large {
            font-family: 'Quicksand', sans-serif;
            font-size: 1.2em;
            font-weight: 700;
            margin: 0;
            color: #4CAF50;
        }

        .header-text-small {
            font-family: 'Quicksand', sans-serif;
            font-size: 0.85em;
            margin: 0;
            color: #757575;
        }

        .header-bottom-info {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-top: 15px;
            padding-top: 10px;
            border-top: 2px solid #e0e0e0;
        }

        .company-placeholder {
            flex-grow: 1;
        }

        .date-info {
            flex-shrink: 0;
            margin-left: 20px;
            text-align: right;
        }

        .sticky-controls-wrapper {
            position: sticky;
            top: 0;
            z-index: 100;
            background-color: #ffffff;
            padding: 15px 30px 10px 30px;
            margin: -30px -30px 0 -30px;
            width: calc(100% + 60px);
            box-sizing: border-box;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.08);
            border-radius: 0 0 12px 12px;
        }

        .actions {
            margin-bottom: 15px;
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }

        .action-button {
            padding: 12px 25px;
            border: none;
            border-radius: 10px;
            font-size: 1em;
            cursor: pointer;
            transition: all 0.2s ease-in-out;
            font-weight: 600;
            position: relative;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1), 0 1px 3px rgba(0, 0, 0, 0.08);
        }

        .action-button::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: linear-gradient(180deg, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0) 50%, rgba(0,0,0,0.05) 100%);
            z-index: 1;
            transition: all 0.2s ease-in-out;
        }

        .action-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 10px rgba(0, 0, 0, 0.15), 0 2px 5px rgba(0, 0, 0, 0.1);
        }

        .action-button:active {
            transform: translateY(1px);
            box-shadow: 0 2px 3px rgba(0, 0, 0, 0.1), 0 1px 2px rgba(0, 0, 0, 0.08);
        }

        .action-button.add {
            background-color: #A5D6A7;
            color: white;
        }
        .action-button.random {
            background-color: #CE93D8;
            color: white;
        }
        .action-button.delete {
            background-color: #EF9A9A;
            color: white;
        }
        .action-button.duplicate {
            background-color: #FFCC80;
            color: #333;
        }
        .action-button.import {
            background-color: #90CAF9;
            color: white;
        }
        .action-button.export {
            background-color: #FFF59D;
            color: #333;
        }
        .action-button.save {
            background-color: #81C784;
            color: white;
        }
        .action-button.load {
            background-color: #64B5F6;
            color: white;
        }
        .action-button.email {
            background-color: #FFB74D;
            color: white;
        }

        .search-filter-controls {
            display: flex;
            gap: 10px;
            margin-bottom: 10px;
            align-items: center;
            flex-wrap: wrap;
        }

        .search-filter-controls input[type="text"],
        .search-filter-controls select {
            padding: 8px 10px;
            border: 1px solid #ccc;
            border-radius: 8px;
            font-size: 0.9em;
            flex-grow: 1;
            max-width: 250px;
        }

        .search-filter-controls button {
            padding: 8px 15px;
            border: none;
            border-radius: 8px;
            background-color: #64B5F6;
            color: white;
            cursor: pointer;
            transition: background-color 0.2s;
        }

        .search-filter-controls button:hover {
            background-color: #42A5F5;
        }

        .table-header-controls {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0;
            padding: 10px 0;
            border-bottom: 1px solid #e0e0e0;
        }

        .table-header-controls h2 {
            margin: 0;
            font-size: 1.2em;
            color: #555;
        }

        .scroll-buttons-wrapper {
            display: flex;
            gap: 5px;
            /* I pulsanti di scorrimento sono fissati rispetto alla viewport così da restare
               sempre visibili mentre si scorre la pagina.  La coordinata 'top'
               viene aggiornata dinamicamente via JavaScript per centrarli
               verticalmente sulla parte visibile della tabella. */
            position: fixed;
            right: 10px;
            top: 50%;
            transform: translateY(-50%);
            z-index: 100;
        }

        .scroll-button {
            padding: 8px 12px;
            background-color: #BBDEFB;
            color: #3F51B5;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 1em;
            font-weight: bold;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            transition: background-color 0.2s, transform 0.2s;
            flex-shrink: 0;
        }

        .scroll-button:hover {
            background-color: #90CAF9;
            transform: translateY(-1px);
        }

        .scroll-button:active {
            transform: translateY(0);
            box-shadow: none;
        }

        .scroll-button:disabled {
            background-color: #e0e0e0;
            color: #a0a0a0;
            cursor: not-allowed;
            box-shadow: none;
        }

        .table-container {
            max-width: 100%;
            overflow: auto;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            height: 500px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 2750px;
            table-layout: fixed;
        }

        table thead {
            z-index: 10;
            background-color: #BBDEFB;
        }

        table thead th {
            position: sticky;
            background-color: #BBDEFB;
            top: 0;
            z-index: 10;
        }

        table thead th:first-child {
            left: 0;
            z-index: 14;
        }
        
        table thead th.col-codice {
            left: 30px;
            z-index: 13;
        }

        table th,
        table td {
            padding: 4px 3px;
            text-align: center;
            border-bottom: 1px solid #eee;
            white-space: normal;
            vertical-align: middle;
            box-sizing: border-box;
        }
        
        .col-cliente, .col-note {
            text-align: left;
        }
        table td .notes-input {
            white-space: normal;
        }

        table th:first-child, table td:first-child { width: 30px; }
        .col-codice { width: 100px; }
        .col-prodotto { width: 280px; }
        .col-cliente { width: 150px; }
        .col-qty-richiesta { width: 130px; }
        .col-giacenza { width: 130px; }
        .col-qty-da-produrre { width: 130px; }
        .col-materie-prime { width: 110px; }
        .col-macchinari { width: 160px; }
        .col-operatore { width: 120px; }
        .col-confez-pezzi { width: 40px; }
        .col-confez-kg-pezzo { width: 50px; }
        .col-prod-data { width: 140px; }
        .col-giorni-produzione { width: 80px; }
        .col-data-confez { width: 140px; }
        .col-cod-confez { width: 140px; }
        .col-lotto-sc { width: 100px; }
        .col-materiale-confez { width: 140px; }
        .col-data-sped { width: 130px; }
        .col-note { width: 300px; }

        table tbody td:first-child {
            position: sticky;
            left: 0;
            background-color: white;
            z-index: 4;
        }

        table tbody td.col-codice {
            position: sticky;
            left: 30px;
            background-color: white;
            z-index: 3;
        }

        table td input[type="text"],
        table td input[type="number"],
        table td select {
            padding: 6px;
            border: 1px solid #cce7f0;
            border-radius: 5px;
            box-sizing: border-box;
            font-size: 0.85em;
            color: #333;
            transition: border-color 0.2s;
            text-align: center;
            width: 100%;
            white-space: normal;
        }

        table td .code-input,
        table td .product-input {
            text-align: center;
        }

        table td input.notes-input {
            height: 50px;
            resize: vertical;
            text-align: left;
            white-space: normal;
        }

        table td input.datepicker {
            width: 100%;
            font-size: 0.8em;
        }
        table td input.packaging-code-input {
            width: 100%;
            font-size: 0.8em;
        }

        .input-with-unit {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 2px;
            width: 100%;
        }
        .input-with-unit input[type="number"],
        .input-with-unit input[type="text"] {
            flex-grow: 1;
            width: auto;
        }
        .unit-label, .unit-select {
            flex-shrink: 0;
            font-size: 0.65em;
            color: #666;
            margin-left: 0;
        }
        .unit-select {
            min-width: 15px;
            width: 30px;
            padding: 0 0px;
            height: 28px;
            font-size: 0.6em;
            text-align-last: center;
        }

        table td input[type="text"]:focus,
        table td input[type="number"]:focus,
        table td select:focus {
            border-color: #007bff;
            outline: none;
            box-shadow: 0 0 5px rgba(0, 123, 255, 0.2);
        }

        table td input[type="checkbox"] {
            transform: scale(1.2);
            cursor: pointer;
        }

        .si-no-select {
            width: 100%;
            padding: 6px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 0.85em;
            background-color: white;
            appearance: none;
            background-image: url('data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%23666%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-6.5%200-12.3%203.4-15.5%208.8-3.2%205.4-3.2%2012%200%2017.4l130%20130c3.2%203.2%207%204.6%2010.8%204.6s7.6-1.4%2010.8-4.6l130-130c3.2-5.4%203.2-12%200-17.4z%22%2F%3E%3C%2Fsvg%3E');
            background-repeat: no-repeat;
            background-position: right 8px center;
            background-size: 12px;
            cursor: pointer;
            text-align-last: center;
        }
        .si-no-select.si {
            background-color: #c8e6c9;
            color: #2e7d32;
        }
        .si-no-select.no {
            background-color: #ffcdd2;
            color: #d32f2f;
        }

        .production-flag {
            font-size: 1.1em;
            vertical-align: middle;
            font-weight: bold;
            flex-shrink: 0;
            margin-left: 0;
            padding-left: 5px;
            width: 20px;
            text-align: center;
            color: transparent;
            transition: color 0.2s ease-in-out;
        }
        .production-flag.visible {
            color: #28a745;
        }

        .validation-feedback {
            font-size: 0.9em;
            margin-left: 5px;
            cursor: help;
            display: inline-block;
            line-height: 1;
            color: transparent;
        }
        .validation-warning-icon {
            color: transparent;
        }
        .validation-error-icon {
            color: transparent;
        }
        .invalid-input-highlight {
            border-color: #cce7f0 !important;
            box-shadow: none !important;
        }

        .gantt-chart-container {
            position: relative;
            background-color: #ffffff;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);
            padding: 30px;
            margin-top: 30px;
            overflow-x: auto;
        }

        .gantt-chart-container h2 {
            font-family: 'Quicksand', sans-serif;
            color: #2c3e50;
            margin-top: 0;
            margin-bottom: 20px;
            font-size: 1.8em;
            border-bottom: 2px solid #e0e0e0;
            padding-bottom: 10px;
        }

        .gantt-chart {
            display: grid;
            grid-template-columns: 200px repeat(14, 1fr);
            border: 2px solid #a0a0a0;
            border-radius: 8px;
            min-width: 1200px;
            gap: 1px;
            background-color: #a0a0a0;
        }
        
        .warehouse-gantt-chart {
            /* Ogni colonna del Gantt di magazzino è larga circa 110px.
               Configuriamo 30 colonne per coprire un arco di un mese.
               La colonna dell'intestazione viene gestita via script e rimane a 200px. */
            grid-template-columns: repeat(30, 110px);
            /* Aggiorniamo la larghezza minima per accomodare 30 colonne (30 × 110px)
               più la colonna intestazione da 200px: totale 3500px.  Lo
               scorrimento orizzontale è gestito dal contenitore. */
            min-width: 3500px;
            margin-top: 15px;
        }

        .gantt-header, .gantt-cell {
            padding: 10px;
            text-align: center;
            font-size: 0.85em;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-start;
            box-sizing: border-box;
            min-height: 70px;
            background-color: white;
        }
        .gantt-header {
            font-weight: bold;
            color: #444;
            position: sticky;
            top: 0;
            z-index: 5;
            background-color: #f0f0f0;
        }
        .gantt-header .day-of-week {
            font-size: 0.75em;
            font-weight: normal;
            color: #777;
        }
        .gantt-header.weekend, .gantt-cell.weekend {
            background-color: #FFECB3;
        }

        .gantt-row-header {
            font-weight: bold;
            text-align: left;
            padding-left: 15px;
            position: sticky;
            left: 0;
            z-index: 4;
            background-color: #f8f8f8;
        }

        .gantt-task {
            /* Rende relativo il contenitore del task per consentire il posizionamento assoluto
               delle icone (lucchetto) e del pallino CQ. */
            position: relative;
            border-radius: 4px;
            margin: 1px;
            color: #1a5276;
            font-weight: 500;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            white-space: normal;
            overflow: hidden;
            text-overflow: ellipsis;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            line-height: 1.1;
            padding: 1px 2px;
            font-size: 0.65em;
            text-align: center;
            width: 99%;
            box-sizing: border-box;
            cursor: help;
        }
        .gantt-task .task-code {
            font-weight: bold;
        }
        .gantt-task .task-details {
            font-size: 0.9em;
        }

        .gantt-task.production-task {
            background-color: #BBDEFB;
        }
        .gantt-task.packaging-task {
            background-color: #C8E6C9;
        }
        .gantt-task.shipping-task {
            background-color: #FFF59D;
            color: #333;
            border: 1px solid #FFD700;
        }

        .gantt-task.production-4xxxx {
            background-color: #64B5F6;
        }
        .gantt-task.packaging-4xxxx {
            background-color: #94E09E;
            color: #333;
        }

        .gantt-task.materie-si {
            border: 2px solid #28a745;
        }
        .gantt-task.materie-no {
            border: 2px solid #dc3545;
        }

        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.5);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            opacity: 0;
            visibility: hidden;
            transition: opacity 0.3s ease, visibility 0.3s ease;
        }

        .modal-overlay.visible {
            opacity: 1;
            visibility: visible;
        }

        .modal-content {
            background-color: #fff;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
            text-align: center;
            max-width: 400px;
            width: 90%;
            transform: translateY(-20px);
            transition: transform 0.3s ease;
        }

        .modal-overlay.visible .modal-content {
            transform: translateY(0);
        }

        .modal-content h3 {
            color: #3F51B5;
            margin-top: 0;
            font-size: 1.5em;
        }

        .modal-content p {
            margin-bottom: 25px;
            line-height: 1.5;
            color: #555;
        }

        .modal-buttons {
            display: flex;
            justify-content: center;
            gap: 15px;
        }

        .modal-button {
            padding: 10px 20px;
            border: none;
            border-radius: 8px;
            font-size: 1em;
            cursor: pointer;
            transition: background-color 0.2s, transform 0.2s;
            font-weight: 600;
        }

        .modal-button.alert, .modal-button.confirm {
            background-color: #4CAF50;
            color: white;
        }

        .modal-button.cancel {
            background-color: #f44336;
            color: white;
        }
        .modal-button.delete {
            background-color: #EF9A9A;
            color: white;
        }

        .modal-button:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }

        .modal-button:active {
            transform: translateY(0);
            box-shadow: none;
        }

        .highlight {
            background-color: yellow;
            font-weight: bold;
        }

        .generic-tooltip {
            position: fixed;
            background-color: rgba(0, 0, 0, 0.85);
            color: white;
            padding: 10px 15px;
            border-radius: 8px;
            font-size: 0.95em;
            z-index: 1001;
            pointer-events: none;
            opacity: 0;
            visibility: hidden;
            transition: opacity 0.05s ease, visibility 0.05s ease;
            max-width: 350px;
            text-align: left;
            line-height: 1.4;
            box-shadow: 0 4px 10px rgba(0,0,0,0.3);
            left: 0;
            top: 0;
        }
        .generic-tooltip.visible {
            opacity: 1;
            visibility: visible;
        }
        .generic-tooltip strong {
            color: #BBDEFB;
        }

        tr.highlight-code-4 td.col-codice input,
        tr.highlight-code-4 td.col-prodotto input,
        tr.highlight-code-4 td.col-cliente input {
            color: red;
            font-weight: bold;
        }

        table th:nth-child(1), table td:nth-child(1) { background-color: #F0F8FF; }
        table th:nth-child(2), table td:nth-child(2) { background-color: #F5FFFA; }
        table th:nth-child(3), table td:nth-child(3) { background-color: #F0FFF0; }
        table th:nth-child(4), table td:nth-child(4) { background-color: #FDF5E6; }
        table th:nth-child(5), table td:nth-child(5) { background-color: #FAEBD7; }
        table th:nth-child(6), table td:nth-child(6) { background-color: #FFF0F5; }
        table th:nth-child(7), table td:nth-child(7) { background-color: #FFE4E1; }
        table th:nth-child(8), table td:nth-child(8) { background-color: #F8F8FF; }
        table th:nth-child(9), table td:nth-child(9) { background-color: #F5F5DC; }
        table th:nth-child(10), table td:nth-child(10) { background-color: #F0FFFF; }
        table th:nth-child(11), table td:nth-child(11) { background-color: #E0F2F7; }
        table th:nth-child(12), table td:nth-child(12) { background-color: #F8E7F0; }
        table th:nth-child(13), table td:nth-child(13) { background-color: #E6F0E6; }
        table th:nth-child(14), table td:nth-child(14) { background-color: #F7F7E0; }
        table th:nth-child(15), table td:nth-child(15) { background-color: #E0E7F7; }
        table th:nth-child(16), table td:nth-child(16) { background-color: #F0E0F7; }
        table th:nth-child(17), table td:nth-child(17) { background-color: #E7F7F0; }
        table th:nth-child(18), table td:nth-child(18) { background-color: #F7F0E0; }
        table th:nth-child(19), table td:nth-child(19) { background-color: #E0F7F0; }
        table th:nth-child(20), table td:nth-child(20) { background-color: #F0F7E0; }

        table tbody td:first-child,
        table tbody td.col-codice {
            background-color: white;
        }
        table thead th {
            background-color: #BBDEFB;
        }

        .daily-production-container {
            background-color: #ffffff;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);
            padding: 30px;
            margin-top: 30px;
            overflow-x: auto;
        }

        .daily-production-container h2 {
            font-family: 'Quicksand', sans-serif;
            color: #2c3e50;
            margin-top: 0;
            margin-bottom: 20px;
            font-size: 1.8em;
            border-bottom: 2px solid #e0e0e0;
            padding-bottom: 10px;
        }

        .daily-production-controls {
            display: flex;
            gap: 10px;
            margin-bottom: 15px;
            align-items: center;
            flex-wrap: wrap;
        }

        .daily-production-controls input[type="text"],
        .daily-production-controls select {
            padding: 8px 10px;
            border: 1px solid #ccc;
            border-radius: 8px;
            font-size: 0.9em;
            flex-grow: 1;
            max-width: 200px;
        }

        .daily-production-controls button {
            padding: 8px 15px;
            border: none;
            border-radius: 8px;
            background-color: #64B5F6;
            color: white;
            cursor: pointer;
            transition: background-color 0.2s;
        }

        .daily-production-controls button:hover {
            background-color: #42A5F5;
        }

        .daily-production-table-wrapper {
            max-width: 100%;
            overflow: auto;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            max-height: 700px;
        }

        .daily-production-table {
            width: 100%;
            border-collapse: collapse;
            min-width: 1800px;
            table-layout: fixed;
            margin-bottom: 20px;
        }
        .daily-production-table th,
        .daily-production-table td {
            padding: 6px 4px;
            text-align: center;
            border-bottom: 1px solid #eee;
            white-space: normal;
            vertical-align: middle;
            box-sizing: border-box;
            /* Aumentiamo la dimensione del font per migliorare
               la leggibilità delle righe della produzione.  Questo rende
               la tabella dei dispositivi medici coerente con le altre
               tabelle operative (spedizioni, arrivi) che utilizzano
               caratteri più grandi. */
            font-size: 0.95em;
        }

        .daily-production-table thead th {
            background-color: #B2EBF2;
            font-weight: bold;
            color: #333;
            position: sticky;
            top: 0;
            z-index: 5;
        }

        .daily-production-table tbody tr.production-row-bg {
            background-color: #F8F8FF;
        }
        .daily-production-table tbody tr.packaging-row-bg {
            background-color: #F0FFF0;
        }

        .daily-production-table tbody tr:hover {
            background-color: #e0f7fa;
        }
        
        .daily-production-table tbody tr.production-4xxxx-bg {
            background-color: #64B5F6 !important;
            color: white;
        }

        .daily-production-table tbody tr.packaging-4xxxx-bg {
            background-color: #94E09E !important;
        }
        .daily-production-table th:first-child, .daily-production-table td:first-child { width: 30px; }
        .daily-production-table .col-daily-codice { width: 70px; }
        .daily-production-table .col-daily-prodotto { width: 180px; }
        /* La colonna cliente è stata ampliata */
        .daily-production-table .col-daily-cliente { width: 140px; }
        .daily-production-table .col-daily-quantita { width: 70px; }
        .daily-production-table .col-daily-macchinario { width: 140px; text-align: center; }
        /* Ridimensioniamo la colonna quantità confezionamento */
        .daily-production-table .col-daily-quantita-confez { width: 110px; }
        .daily-production-table .col-daily-operazioni { width: 300px; }
        /* Allarga leggermente la colonna Operatore e restringe la colonna Data Avallo per migliorare la leggibilità */
        .daily-production-table .col-daily-operatori { width: 150px; }
        .daily-production-table .col-daily-esito { width: 80px; }
        .daily-production-table .col-daily-qty-prodotta { width: 100px; }
        .daily-production-table .col-daily-lotto { width: 90px; }
        .col-daily-tu { width: 70px; }
        .col-daily-ts { width: 70px; }
        .col-daily-data-avallo { width: 80px; }

        /* Nuove colonne per la tabella giornaliera: OPE (Ordine di Produzione Esterno) e OV (Ordine di Vendita).
           Queste colonne sostituiscono le precedenti colonne TU/TS e sono posizionate all'inizio della tabella
           giornaliera.  Manteniamo una larghezza fissa per evitare che si restringano troppo. */
        /* Colonna OP (ex OPE): larghezza fissa per il numero d'ordine di produzione */
        .col-daily-op { width: 70px; }
        .col-daily-ov  { width: 70px; }

        .daily-production-table input[type="text"],
        .daily-production-table input[type="number"],
        .daily-production-table select {
            padding: 5px;
            border: 1px solid #cce7f0;
            border-radius: 4px;
            box-sizing: border-box;
            font-size: 0.8em;
            width: 100%;
            text-align: center;
            white-space: normal;
        }
        .daily-production-table .col-daily-operazioni select {
            text-align: left;
            height: auto;
            min-height: 28px;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        .daily-production-table .col-daily-operatori input { text-align: left; }
        .daily-production-table .col-daily-macchinario input { text-align: center; }
        .daily-production-table td {
            vertical-align: top;
        }

        .sales-order-container {
            background-color: #ffffff;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);
            padding: 30px;
            margin-top: 30px;
            overflow-x: auto;
        }

        .sales-order-container h2 {
            font-family: 'Quicksand', sans-serif;
            color: #2c3e50;
            margin-top: 0;
            margin-bottom: 20px;
            font-size: 1.8em;
            border-bottom: 2px solid #e0e0e0;
            padding-bottom: 10px;
        }

        .sales-order-controls {
            display: flex;
            gap: 10px;
            margin-bottom: 15px;
            align-items: center;
            flex-wrap: wrap;
        }

        .sales-order-controls button {
            padding: 8px 15px;
            border: none;
            border-radius: 8px;
            background-color: #64B5F6;
            color: white;
            cursor: pointer;
            transition: background-color 0.2s;
        }

        .sales-order-controls button:hover { background-color: #42A5F5; }
        .sales-order-table-wrapper {
            max-width: 100%;
            overflow: auto;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            max-height: 400px;
        }

        .sales-order-table {
            width: 100%;
            border-collapse: collapse;
            min-width: 1200px;
            table-layout: fixed;
        }

        .sales-order-table th,
        .sales-order-table td {
            padding: 6px 4px;
            text-align: center;
            border-bottom: 1px solid #eee;
            white-space: normal;
            vertical-align: middle;
            box-sizing: border-box;
            font-size: 0.95em;
        }

        .sales-order-table thead th {
            background-color: #CFD8DC;
            font-weight: bold;
            color: #333;
            position: sticky;
            top: 0;
            z-index: 5;
        }

        .sales-order-table .col-ov-flag { width: 60px; position: relative; }
        .sales-order-table th:first-child, .sales-order-table td:first-child { width: 30px; }
        .sales-order-table .col-ov { width: 80px; }
        .sales-order-table .col-ov-codice { width: 100px; }
        .sales-order-table .col-ov-descrizione { width: 250px; text-align: left;}
        .sales-order-table .col-ov-quantita { width: 100px; }
        .sales-order-table .col-ov-um { width: 60px; }
        .sales-order-table .col-ov-data-consegna { width: 120px; }
        .sales-order-table .col-ov-data-richiesta-cliente { width: 120px; }
        .sales-order-table .col-ov-data-conferma { width: 120px; }
        .sales-order-table .col-ov-note { width: 300px; text-align: left; }

        .sales-order-table input[type="text"],
        .sales-order-table input[type="number"],
        .sales-order-table select {
            padding: 5px;
            border: 1px solid #cce7f0;
            border-radius: 4px;
            box-sizing: border-box;
            font-size: 0.95em;
            width: 100%;
            text-align: center;
            white-space: normal;
        }
        .sales-order-table .col-ov-descrizione input,
        .sales-order-table .col-ov-note input {
            text-align: left;
        }

        .ov-flag-icon {
            font-size: 2.2em;
            line-height: 1;
            vertical-align: middle;
            display: inline-block;
            width: 1.2em;
            text-align: center;
            cursor: help;
        }
        .ov-flag-icon.red-triangle { color: red; }
        .ov-flag-icon.yellow-triangle { color: #FFD700; }
        .ov-flag-icon.dark-yellow-square { color: #DAA520; }
        .ov-flag-icon.green-square { color: green; }

        /* Ridimensiona le icone dei flag OV affinché non siano troppo invasive. Riduciamo
           la dimensione del simbolo e dell'eventuale punto esclamativo, mantenendolo
           comunque leggibile. */
        .ov-flag-icon {
            /* Ingrandiamo le icone dei flag OV per renderle più visibili,
               uniformando la loro altezza a quella della riga. */
            font-size: 1.1em;
            line-height: 1;
        }
        .ov-flag-icon .exclamation-mark {
            font-size: 0.75em;
            top: -0.15em;
            left: -0.4em;
        }

        /* Separatore grafico tra la legenda CQ e la legenda QA nella sezione delle spedizioni */
        .legend-separator {
            /* Un ampio spazio verticale tra le legende CQ e QA per evitare che sembrino
               un unico blocco. Incrementiamo l'altezza a 36px per una chiara separazione. */
            height: 36px;
            width: 100%;
        }

        .ov-flag-icon .exclamation-mark {
            position: relative;
            top: -0.2em; left: -0.5em;
            font-size: 0.8em;
            color: black;
            font-weight: bold;
        }
        .header-section {
            background-color: #f0f0f0;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .file-status {
            font-weight: bold;
        }
        .file-status .green {
            color: green;
        }
        .file-status .red {
            color: red;
        }
        /* =================================================================== */
/* ==> INCOLLA QUESTO BLOCCO PER IL CONTROLLO TOTALE DELLE COLONNE <== */
/* =================================================================== */

#analisiTable {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
    table-layout: fixed; /* FONDAMENTALE: Obbliga la tabella a rispettare le larghezze definite */
}

/* ===== ADD: Modali scorrevoli e blocco background scroll ===== */
/* Imposta un'altezza massima ai contenuti dei modali affinché il contenuto
   possa scorrere internamente se supera l'80% dell'altezza della finestra. */
.modal-content {
  max-height: 80vh;
  overflow-y: auto;
  -webkit-overflow-scrolling: touch; /* scroll fluido su dispositivi iOS */
}
/* Quando un modale è aperto, impedisci lo scroll della pagina di sfondo */
body.modal-open {
  overflow: hidden;
}
/* Impedisci che lo scroll passi attraverso il modale (overscroll) */
.modal-overlay.visible {
  overscroll-behavior: contain;
}

/* ----------------------------------------------------------------------------
 * Ghost horizontal scrollbars
 * Per le tabelle con molti campi (come i programmi giornalieri di spedizione e
 * arrivo), forniamo una barra di scorrimento orizzontale "fantasma" che
 * rimane centrata verticalmente sulla porzione visibile della tabella.  In
 * questo modo l'utente può scorrere lateralmente anche quando la barra di
 * scorrimento nativa è fuori dal campo visivo.  La barra fantasma replica il
 * movimento della barra originale e viceversa.
 * ------------------------------------------------------------------------- */
.ghost-scrollbar {
    position: fixed;
    height: 12px;
    overflow-x: auto;
    overflow-y: hidden;
    background: transparent;
    z-index: 120;
    display: none; /* viene mostrata dinamicamente quando la tabella è visibile */
}
.ghost-scrollbar::-webkit-scrollbar {
    height: 8px;
}
.ghost-scrollbar::-webkit-scrollbar-track {
    background: rgba(0, 0, 0, 0.05);
}
.ghost-scrollbar::-webkit-scrollbar-thumb {
    background: rgba(0, 0, 0, 0.3);
    border-radius: 4px;
}

/* Barre di scorrimento orizzontale ancorate al Gantt (sopra e sotto) */
.gantt-inline-scrollbar {
  height: 12px;
  overflow-x: auto;
  overflow-y: hidden;
  background: transparent;
  margin: 6px 0;            /* attacca la barra alla griglia senza distacchi eccessivi */
}
.gantt-inline-scrollbar::-webkit-scrollbar { height: 10px; }
.gantt-inline-scrollbar::-webkit-scrollbar-track { background: rgba(0,0,0,0.05); }
.gantt-inline-scrollbar::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.3); border-radius: 4px; }


/* ====== BARRE DI SCORRIMENTO ANCORATE (SOPRA e SOTTO) ====== */
.dock-scrollbar {
  height: 12px;
  overflow-x: auto;
  overflow-y: hidden;
  background: transparent;
  z-index: 6;
}

/*
 * Miglioramento usabilità: mantiene fissi i giorni della settimana nel grafico
 * di Gantt delle spedizioni/arrivi anche durante lo scorrimento verticale.
 * Il wrapper del grafico ora gestisce lo scroll verticale in autonomia, per
 * evitare che l'intestazione dei giorni scompaia quando si visualizzano
 * molte righe di ordini.  Impostando overflow-y su auto viene abilitato
 * lo scroll verticale all'interno del contenitore, e grazie a position: sticky
 * sulle celle di intestazione, i giorni restano visibili in cima.
 */
#warehouseGanttScrollWrapper {
  overflow-y: auto;
  /* Limita l'altezza del grafico in modo da far comparire la barra di scorrimento
     verticale quando il contenuto supera tale altezza.  L'altezza scelta
     (60vh) occupa circa il 60% dell'area visibile della finestra, ma può
     essere modificata in base alle esigenze di layout. */
  max-height: 60vh;
}

/* Le celle di intestazione (giorni) nel Gantt rimangono appiccicate in alto
   durante lo scorrimento verticale all'interno del wrapper.  Lo z-index
   garantisce che rimangano sopra le altre celle. */
#warehouseGanttScrollWrapper .gantt-header {
  position: sticky;
  top: 0;
  z-index: 10;
}
.dock-scrollbar::-webkit-scrollbar { height: 10px; }
.dock-scrollbar::-webkit-scrollbar-track { background: rgba(0,0,0,0.05); }
.dock-scrollbar::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.30); border-radius: 4px; }

.dock-scrollbar.top {
  position: sticky;
  top: 0;
  margin-bottom: 4px;
  background: linear-gradient(to bottom, rgba(255,255,255,1), rgba(255,255,255,0.90));
}

.dock-scrollbar.bottom {
  position: sticky;
  bottom: 0;
  margin-top: 4px;
  background: linear-gradient(to top, rgba(255,255,255,1), rgba(255,255,255,0.90));
}

/* I contenitori che hanno le barre ancorate aggiungono un po' di padding
   per evitare sovrapposizioni con intestazioni/corpo. */
.has-dock-scrollbars { padding-top: 14px; padding-bottom: 14px; }

/* Mantiene l'intestazione della tabella sotto la barra superiore,
   SOLO per i wrapper che hanno le barre ancorate. */
.has-dock-scrollbars .daily-production-table thead th {
  top: 12px !important;
}

/* I wrapper devono essere il contesto di scorrimento per sticky */
.daily-production-table-wrapper,
#warehouseGanttChartContainer {
  position: relative;
}

        /* ------------------------------------------------------------------
         * Barre di scorrimento orizzontali aggiuntive per il Gantt Spedizioni.
         * La barra esterna viene posizionata tra la tabella delle spedizioni
         * giornaliere e il Gantt, in modo che l'utente possa scorrere
         * orizzontalmente il grafico anche quando si trova a metà della pagina.
         * Viene sincronizzata via script con lo scorrimento del contenitore
         * del Gantt.
         * ------------------------------------------------------------------ */
        .gantt-external-scrollbar {
          height: 12px;
          overflow-x: auto;
          overflow-y: hidden;
          background: transparent;
          margin: 6px 0;
        }
        .gantt-external-scrollbar::-webkit-scrollbar {
          height: 10px;
        }
        .gantt-external-scrollbar::-webkit-scrollbar-track {
          background: rgba(0,0,0,0.05);
        }
        .gantt-external-scrollbar::-webkit-scrollbar-thumb {
          background: rgba(0,0,0,0.3);
          border-radius: 4px;
        }

        /* Pulsanti laterali per lo scroll del Gantt Spedizioni */
        .gantt-scroll-buttons-wrapper {
          position: fixed;
          right: 10px;
          top: 50%;
          transform: translateY(-50%);
          z-index: 110;
          display: flex;
          gap: 5px;
        }

/* ------------------------------------------------------------------ */
/* Styles for the Packing List modal and its contents. These styles   */
/* create a centered modal with a scrollable list of orders and       */
/* products, plus clearly styled buttons for creating and closing     */
/* the packing list. They are loaded after the ghost scrollbar        */
/* definitions so they don't interfere with existing styles.          */
/* ------------------------------------------------------------------ */
.packing-list-modal {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.5);
    display: none;
    justify-content: center;
    align-items: center;
    z-index: 2000;
}
.packing-list-content {
    background-color: #ffffff;
    padding: 20px;
    border-radius: 8px;
    max-width: 800px;
    width: 90%;
    max-height: 80vh;
    display: flex;
    flex-direction: column;
}
.packing-list-content h3 {
    margin-top: 0;
    margin-bottom: 15px;
    font-size: 1.4em;
    color: #333333;
}
.packing-list-items {
    list-style: none;
    padding-left: 0;
    margin: 0;
    flex: 1 1 auto;
    overflow-y: auto;
    border-top: 1px solid #eee;
    border-bottom: 1px solid #eee;
}
.packing-list-item {
    padding: 8px 0;
    border-bottom: 1px solid #f0f0f0;
}
.packing-list-item:last-child {
    border-bottom: none;
}
.packing-list-item > label {
    font-weight: bold;
    cursor: pointer;
}
.packing-list-item ul {
    list-style: none;
    margin-left: 20px;
    padding-left: 0;
    margin-top: 5px;
}
.packing-list-item li {
    margin-bottom: 4px;
}
.packing-list-modal-footer {
    display: flex;
    justify-content: flex-end;
    gap: 10px;
    margin-top: 15px;
}
.packing-list-modal-footer button {
    padding: 8px 16px;
    border: none;
    border-radius: 6px;
    cursor: pointer;
    font-size: 0.9em;
    font-weight: 600;
}
#packingListCreateBtn {
    background-color: #4CAF50;
    color: white;
}
#packingListCloseBtn {
    background-color: #f44336;
    color: white;
}

#analisiTable th,
#analisiTable td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: center;      /* Centra il testo in tutte le celle */
    vertical-align: middle;  /* Allinea verticalmente al centro */
    white-space: normal;      /* Permette al testo di andare a capo */
    word-wrap: break-word;    /* Forza l'andata a capo per parole lunghe */
}

/* --- CONTROLLO MANUALE LARGHEZZA COLONNE --- */
/* Modifichi queste percentuali come preferisce */

#analisiTable th:nth-child(1) { width:  2%; } 
#analisiTable th:nth-child(2) { width: 30%; }
#analisiTable th:nth-child(3) { width:  5%; }
#analisiTable th:nth-child(4) { width:  5%; }
#analisiTable th:nth-child(5) { width:  2.3%; }
/* Le altre colonne di analisi si divideranno lo spazio rimanente. 
   Se vuole impostare una larghezza specifica per una di esse, 
   ad esempio la colonna "Colore" (che è la settima), aggiunga:

   #analisiTable th:nth-child(7) { width: 80px; } 

*/
        /* La parentesi graffa seguente era superflua e causava errori di parsing.
           È stata rimossa per correggere il CSS. */
        #analisiTable th {
            background-color: #f2f2f2;
            text-align: center;
            vertical-align: middle;
        }
        .analisi-table-actions {
            display: flex;
            gap: 5px;
            justify-content: flex-end;
            margin-bottom: 10px;
        }
        .analisi-table-actions button {
            padding: 5px 10px;
            border: none;
            cursor: pointer;
            border-radius: 4px;
        }
        .btn-add {
            background-color: #28a745;
            color: white;
        }
        .btn-duplicate {
            background-color: #ffc107;
            color: #333;
        }
        .btn-delete {
            background-color: #dc3545;
            color: white;
        }
        .analysis-cell {
            text-align: center;
            white-space: nowrap;
        }
        .analysis-cell .analysis-name {
            font-size: 1em; /* leggermente più grande per migliorare la leggibilità */
            font-weight: bold;
            display: block;
            margin-bottom: 5px;
            white-space: normal;
        }
        .analysis-cell .method-in-cell {
            font-size: 0.8em; /* aumenta la dimensione del metodo per visibilità */
            color: #777;
            display: block;
            margin-bottom: 3px;
        }
        
/* VERSIONE AGGIORNATA */
.conformity-box {
    display: inline-block;
    width: 20px;
    height: 20px;
    border: 2px solid #555; /* Bordo più spesso e scuro */
    margin: 2px;
    cursor: pointer;
    vertical-align: middle;
    position: relative;
}

        .conformity-box.green-flag:after {
            content: '✔';
            color: green;
            font-size: 14px;
            line-height: 20px;
            position: absolute;
            top: 0;
            left: 50%;
            transform: translateX(-50%);
        }
        .conformity-box.red-x:after {
            content: '✖';
            color: red;
            font-size: 14px;
            line-height: 20px;
            position: absolute;
            top: 0;
            left: 50%;
            transform: translateX(-50%);
        }
        .collapsible-row {
            cursor: pointer;
        }
        .hidden-row {
            display: none;
        }
        .rotate-icon {
            display: inline-block;
            transition: transform 0.3s ease;
        }
        .rotate-icon.rotated {
            transform: rotate(90deg);
        }
        .method-column {
            font-size: 0.8em;
            color: #555;
            padding: 0 5px;
        }
        .flag {
            color: green;
            font-weight: bold;
            margin-left: 5px;
        }
        .file-status-flag {
            position: absolute;
            top: -10px;
            right: 0;
            font-size: 0.8em;
            color: #28a745;
            display: none;
        }
        @media print {
            body:not(.printing-logbook) #logbookContainer {
                display: none !important;
            }

            body.printing-logbook .sticky-controls-wrapper,
            body.printing-logbook header,
            body.printing-logbook .table-container,
            body.printing-logbook .sales-order-container,
            body.printing-logbook .gantt-chart-container,
            body.printing-logbook .daily-production-container {
                display: none !important;
            }

            body.printing-logbook #logbookContainer {
                display: block !important;
                box-shadow: none;
                border: none;
                padding: 1cm;
                margin: 0;
            }

            body.printing-logbook #logbookContainer h2,
            body.printing-logbook #logbookContainer .daily-production-controls,
            body.printing-logbook #logbookContainer #clearLogbookBtn {
                display: none;
            }
            
            body.printing-logbook #logbookContainer > div {
                max-height: none !important;
                overflow-y: visible !important;
            }

            body.printing-logbook #logbookContent {
                font-size: 9pt;
                white-space: pre-wrap;
                word-wrap: break-word;
            }
#print-header-info {
                display: block !important;
                text-align: center;
                border-bottom: 2px solid #333;
                padding-bottom: 10px;
                margin-bottom: 15px;
            }
            #print-header-info h3 {
                margin: 0;
                font-size: 16pt;
                font-weight: bold;
                color: #000;
            }
            @page {
                size: A4 landscape;
                margin: 1cm;
            }

            /* Le regole @page non possono essere annidate in selettori; rimosse per evitare errori CSS. */


            body { padding: 0; margin: 0; background-color: #fff; }
            .container { box-shadow: none; padding: 0px; width: 100%; max-width: none; }
            .sticky-controls-wrapper,
            .header-layout-table,
            .header-bottom-info,
            .info-note,
            .gantt-chart-container,
            .table-container,
            .sales-order-container,
            .warehouse-gantt-chart-container {
                display: none !important;
            }
            .daily-production-container {
                display: block !important;
                box-shadow: none;
                padding: 0;
                margin: 0;
                overflow: visible !important;
                page-break-inside: avoid !important;
            }
            .daily-production-controls { display: none; }
            .daily-production-table-wrapper { overflow: visible !important; height: auto !important; max-height: none !important; border: none; box-shadow: none; }
            .daily-production-table { width: 100%; min-width: unset; table-layout: fixed; }
            .daily-production-table th, .daily-production-table td {
                border: 1px solid #ccc;
                padding: 4px 2px;
                text-align: center;
                border-bottom: 1px solid #eee;
                white-space: normal;
                word-wrap: break-word;
                vertical-align: middle;
                box-sizing: border-box;
                font-size: 7pt;
                font-weight: bold;
                line-height: 1.2;
            }
            /* Allinea al centro il contenuto di tutte le celle della tabella giornaliera nella stampa */
            .daily-production-table .col-daily-prodotto,
            .daily-production-table .col-daily-cliente,
            .daily-production-table .col-daily-operazioni,
            .daily-production-table .col-daily-operatori,
            .daily-production-table .col-daily-lotto { text-align: center; }
            .daily-production-table thead { display: table-header-group; }
            .daily-production-table thead th { background-color: #e0e0e0 !important; -webkit-print-color-adjust: exact; color-adjust: exact; position: static; }
            .daily-production-table tbody tr { page-break-inside: avoid; }

            /* Centra il contenuto di tutte le celle (intestazioni e corpo) nella tabella giornaliera in stampa */
            .daily-production-table th,
            .daily-production-table td {
                text-align: center;
            }
            .daily-production-table input[type="text"],
            .daily-production-table input[type="number"],
            .daily-production-table select { border: none; background-color: transparent; padding: 0; font-size: inherit; width: 100%; text-align: inherit; box-shadow: none; }
            .daily-production-table th:first-child, .daily-production-table td:first-child { width: 1.5%; }
            /* Specifica le larghezze per le colonne principali nel programma giornaliero in stampa.
               OP e OV sono dimensionate per contenere 5 cifre ciascuno.  Codice è ampliato per 9 caratteri.
               Prodotto e Cliente sono leggermente ridotti per equilibrare lo spazio. */
            .daily-production-table .col-daily-op { width: 4%; }
            .daily-production-table .col-daily-ov { width: 4%; }
            .daily-production-table .col-daily-codice { width: 6%; }
            .daily-production-table .col-daily-prodotto { width: 11%; }
            .daily-production-table .col-daily-cliente { width: 10%; }
            .daily-production-table .col-daily-lotto { width: 4%; }
            .daily-production-table .col-daily-quantita { width: 5%; }
            .daily-production-table .col-daily-macchinario { width: 8%; }
            .daily-production-table .col-daily-quantita-confez { width: 4.5%; }
            .daily-production-table .col-daily-operazioni { width: 19%; }
            .daily-production-table .col-daily-operatori { width: 6%; }
            .daily-production-table .col-daily-esito { width: 3%; }
            .daily-production-table .col-daily-qty-prodotta { width: 6%; }
            .daily-production-table .col-daily-tu { width: 2%; }
            .col-daily-ts { width: 2%; }
            /* La colonna Data Avallo viene allargata per contenere due righe */
            .col-daily-data-avallo { width: 4%; }

            .daily-production-table tbody tr.production-row-bg { background-color: #F8F8FF !important; -webkit-print-color-adjust: exact; color-adjust: exact; }
            .daily-production-table tbody tr.packaging-row-bg { background-color: #F0FFF0 !important; -webkit-print-color-adjust: exact; color-adjust: exact; }
        }
/* Aggiungi questo codice alla fine della sezione <style> */
#analisiTable .analysis-cell {
    vertical-align: middle; /* Centra verticalmente */
    text-align: center;      /* Centra orizzontalmente */
}
/* Aggiungi anche questo codice alla fine della sezione <style> */
#analisiTable th .analysis-name {
    white-space: normal; /* Permette al testo di andare a capo */
    max-width: 120px;    /* Imposta una larghezza massima per la colonna */
    margin: 0 auto;      /* Centra il testo nella cella */
}
/* NUOVI STILI PER LA SEZIONE ANALISI */
#analisiTable .analysis-cell {
    vertical-align: middle;
    text-align: center;
}
#analisiTable th .analysis-name {
    white-space: normal;
    max-width: 150px; /* Larghezza massima colonna analisi */
    margin: 0 auto;
}
.status-indicator {
    display: inline-block;
    width: 18px;
    height: 18px;
    border-radius: 50%;
    border: 1px solid #ccc;
}
.status-red { background-color: #f44336; }
.status-yellow { background-color: #ffeb3b; }
.status-green { background-color: #4CAF50; }
.status-nc {
    font-weight: bold;
    font-size: 1.5em;
    color: #f44336;
    line-height: 1;
}
/* NUOVI STILI PER STAMPA ANALISI E STATO */
body.printing-analisi .sticky-controls-wrapper,
body.printing-analisi header,
body.printing-analisi .table-container,
body.printing-analisi .sales-order-container,
body.printing-analisi .gantt-chart-container,
body.printing-analisi .daily-production-container:not(#analisiContainer),
body.printing-analisi #logbookContainer {
    display: none !important;
}

body.printing-analisi #analisiContainer {
    display: block !important;
    box-shadow: none; border: none; padding: 0; margin: 0;
}

body.printing-analisi #analisiContainer .analisi-table-actions,
body.printing-analisi #analisiContainer .daily-production-controls {
    display: none;
}

@media print {
    /* Disegna una linea nera tra le righe del programma giornaliero per
       evidenziare la separazione tra una riga e l'altra.  Tutti i
       valori sono inoltre centrati verticalmente e orizzontalmente per
       facilitare la compilazione a mano. */
    #dailyProductionTable tbody tr {
        border-bottom: 1px solid #000 !important;
    }
    #dailyProductionTable th, #dailyProductionTable td {
        text-align: center !important;
        vertical-align: middle !important;
    }
    body:not(.printing-analisi) #analisiContainer {
        display: none !important;
    }
}

.status-indicator {
    display: inline-block;
    width: 18px;
    height: 18px;
    border-radius: 50%;
    border: 1px solid #ccc;
}
.status-red { background-color: #f44336; }
.status-yellow { background-color: #ffeb3b; }
.status-green { background-color: #4CAF50; }
.status-nc {
    font-weight: bold;
    font-size: 1.5em;
    color: #f44336;
    line-height: 1;
}

/* Migliora la leggibilità della tabella analisi aumentando leggermente la dimensione del testo */
#analisiTable th,
#analisiTable td {
    font-size: 0.95em;
}

/* Imposta una larghezza più contenuta per la colonna Prodotto migliorando la leggibilità */
#analisiTable th:nth-child(2),
#analisiTable td:nth-child(2) {
    width: 350px; /* Ridotta la larghezza per evitare colonne troppo ampie */
    min-width: 350px;
    white-space: normal;
}

/* Stile per la TEXTAREA per renderla adattabile e senza bordi */
#analisiTable td:nth-child(2) textarea.analisi-input {
    width: 100%;
    border: none;
    background-color: transparent;
    resize: vertical; /* Permette il ridimensionamento verticale */
    min-height: 40px; /* Altezza minima iniziale */
    font-size: inherit;
    font-family: inherit;
    font-weight: inherit;
    color: inherit;
    padding: 2px;
    box-sizing: border-box;
}

/* Corregge il layout della riga principale per far espandere la textarea */
#analisiTable .collapsible-row td:nth-child(2) strong {
    flex-grow: 1;
    display: block;
}


/* Stile per la riga principale della tabella Analisi (come da foto) */
#analisiTable .collapsible-row td:nth-child(2) {
    display: flex;
    align-items: center; /* Allinea verticalmente freccia e testo */
    text-align: left;
}
#analisiTable .collapsible-row strong .analisi-input {
    font-family: 'Georgia', 'Times New Roman', serif; /* Cambia il font */
    font-weight: bold;
    color: #000; /* Testo più scuro */
    font-size: 0.9em; /* Testo leggermente più grande */
}

/* =================================================================== */
/* ==> NUOVI STILI PER ICONA APRI/CHIUDI ANALISI (QUADRATA) <== */
/* =================================================================== */

#analisiTable .rotate-icon {
    font-size: 18px;          /* Dimensione del simbolo "+" */
    font-weight: bold;        /* Rende il simbolo più marcato */
    color: #000000;           /* MODIFICA: Nero pieno per il simbolo */
    
    /* Creazione del quadrato */
    border: 1px solid #bbbbbb;/* Bordo grigio leggermente più scuro */
    border-radius: 4px;       /* MODIFICA: Angoli leggermente arrotondati (non più un cerchio) */
    width: 22px;              /* Larghezza fissa */
    height: 22px;             /* Altezza fissa */

    /* Centratura perfetta del simbolo */
    display: flex;
    align-items: center;
    justify-content: center;
    line-height: 1;

    /* Aspetto e usabilità */
    cursor: pointer;
    margin-right: 12px;
    transition: background-color 0.2s ease;
}

#analisiTable .rotate-icon:hover {
    background-color: #f0f0f0; /* Sfondo al passaggio del mouse */
}



/* === NUOVI STILI TOOLTIP e ICONA INFO === */
.tooltip-container {
    display: flex;
    gap: 8px;
    padding: 5px;
}
.tooltip-box {
    padding: 10px 14px;
    border-radius: 8px;
    border: 2px solid;
    min-width: 280px;
    font-size: 0.9em;
    background-color: #fff;
}
.tooltip-box h3 {
    margin: 0 0 8px 0;
    padding-bottom: 5px;
    border-bottom: 1px solid;
    font-size: 1.1em;
}
.tooltip-box p { margin: 0; line-height: 1.6; }
.tooltip-box strong { font-weight: 700; }

/* Stile specifico per il box di Dettaglio OPI
   - Colore di bordo nero per evidenziarlo
   - Colore del testo nero per migliorare la leggibilità
   - Sfondo bianco (ereditato da .tooltip-box) */
.tooltip-box.opi-info-tooltip {
    border-color: #000;
    color: #000;
    /* Non impostiamo il font-weight a livello di box per permettere ai valori
       di rimanere con peso normale. Le etichette sono in grassetto grazie al tag <strong>. */
}

/* Le etichette (tag <strong>) all'interno del box OPI devono essere nere per
   garantire visibilità, indipendentemente dai colori ereditati. */
.tooltip-box.opi-info-tooltip strong {
    color: #000;
}
.tooltip-box.opi-info-tooltip h3 {
    border-bottom-color: #000;
    color: #000;
}

.production-tooltip {
    background-color: #E3F2FD;
    border-color: #64B5F6;
    color: #0D47A1;
}
.production-tooltip h3 { border-bottom-color: #90CAF9; color: #1565C0; }
.production-tooltip strong { color: #1976D2; }

.packaging-tooltip {
    background-color: #E8F5E9;
    border-color: #81C784;
    color: #1B5E20;
}
.packaging-tooltip h3 { border-bottom-color: #A5D6A7; color: #2E7D32; }
.packaging-tooltip strong { color: #388E3C; }

.info-icon {
    font-weight: bold;
    cursor: pointer;
    color: #007bff;
    margin-left: 5px;
    font-size: 1.1em;
    display: inline-block;
}

/* === NUOVI STILI TOOLTIP SPEDIZIONE === */
.shipping-info-tooltip {
    background-color: #FFF9C4; /* Giallo chiaro */
    border-color: #FFEB3B;
    color: #795548;
}
.shipping-info-tooltip h3 { border-bottom-color: #FFF59D; color: #5D4037; }
.shipping-info-tooltip strong { color: #BF360C; }

.shipping-contact-tooltip {
    background-color: #F1F8E9; /* Verde molto chiaro */
    border-color: #AED581;
    color: #33691E;
}
.shipping-contact-tooltip h3 { border-bottom-color: #C5E1A5; color: #558B2F; }
.shipping-contact-tooltip strong { color: #689F38; }



/* --- 1. Stili FONDAMENTALI Comuni per Entrambe le Tabelle --- */
#shippingScheduleTable,
#arrivalScheduleTable {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed !important; /* Forza il rispetto delle larghezze definite */
}

#shippingScheduleTable th, #shippingScheduleTable td,
#arrivalScheduleTable th, #arrivalScheduleTable td {
    padding: 4px 3px;
    text-align: center;
    border-bottom: 1px solid #eee;
    white-space: normal !important;   /* FONDAMENTALE per mandare il testo a capo */
    word-wrap: break-word !important; /* FONDAMENTALE per parole lunghe */
    vertical-align: middle;
    box-sizing: border-box;
}

#shippingScheduleTable th,
#arrivalScheduleTable th {
    background-color: #CFD8DC;
    font-weight: bold;
    color: #333;
    position: sticky;
    top: 0;
    z-index: 5;
}

#shippingScheduleTable input, #shippingScheduleTable select,
#arrivalScheduleTable input, #arrivalScheduleTable select {
    padding: 6px;
    border: 1px solid #cce7f0;
    border-radius: 5px;
    box-sizing: border-box;
    font-size: 0.85em;
    width: 100%;
    text-align: center;
}

/* Allineamento a sinistra solo per le colonne di testo lunghe */
#shippingScheduleTable td:nth-child(4) input,  /* Descrizione Articolo (Spedizioni) */
#shippingScheduleTable td:nth-child(9) input,  /* Ragione Sociale (Spedizioni) */
#shippingScheduleTable td:nth-child(10) input, /* Rif. Cliente (Spedizioni) */
#shippingScheduleTable td:nth-child(11) input, /* Indirizzo (Spedizioni) */
#arrivalScheduleTable td:nth-child(4) input,   /* Descrizione Articolo (Arrivi) */
#arrivalScheduleTable td:nth-child(10) input,  /* Ragione Sociale (Arrivi) */
#arrivalScheduleTable td:nth-child(11) input,  /* Rif. Cliente (Arrivi) */
#arrivalScheduleTable td:nth-child(12) input {  /* Indirizzo (Arrivi) */
    text-align: left;
}

/* --- 2. Larghezze Specifiche e Identiche per le Colonne --- */
#shippingScheduleTable { min-width: 2300px; }
#arrivalScheduleTable { min-width: 2450px; } /* Più larga per la colonna Layout */

/* Colonne Comuni */
#shippingScheduleTable th:nth-child(1), #arrivalScheduleTable th:nth-child(1) { width: 30px; }
#shippingScheduleTable th:nth-child(2), #arrivalScheduleTable th:nth-child(2) { width: 80px; }
#shippingScheduleTable th:nth-child(3), #arrivalScheduleTable th:nth-child(3) { width: 100px; }
#shippingScheduleTable th:nth-child(4), #arrivalScheduleTable th:nth-child(4) { width: 430px; }

/* Colonna Layout (solo Tabella Arrivi) */
#arrivalScheduleTable th:nth-child(5) { width: 150px; }

/* Colonne restanti (con indici sfalsati per la tabella Arrivi) */
#shippingScheduleTable th:nth-child(5), #arrivalScheduleTable th:nth-child(6) { width: 90px; }
#shippingScheduleTable th:nth-child(6), #arrivalScheduleTable th:nth-child(7) { width: 50px; }
#shippingScheduleTable th:nth-child(7), #arrivalScheduleTable th:nth-child(8) { width: 90px; }
#shippingScheduleTable th:nth-child(8), #arrivalScheduleTable th:nth-child(9) { width: 90px; }
#shippingScheduleTable th:nth-child(9), #arrivalScheduleTable th:nth-child(10){ width: 300px; }
#shippingScheduleTable th:nth-child(10),#arrivalScheduleTable th:nth-child(11){ width: 200px; }
#shippingScheduleTable th:nth-child(11),#arrivalScheduleTable th:nth-child(12){ width: 320px; }
#shippingScheduleTable th:nth-child(12),#arrivalScheduleTable th:nth-child(13){ width: 70px; }
#shippingScheduleTable th:nth-child(13),#arrivalScheduleTable th:nth-child(14){ width: 140px; }

/* -------------------------------------------------------------------
 * Stili per la tabella "Merce in Scadenza".  Questa tabella replica
 * le caratteristiche di impaginazione delle altre tabelle di
 * spedizione/arrivo, inclusi larghezze fisse, sfondo sticky e input
 * allineati.  Le dimensioni sono adattate alle colonne specifiche
 * (codice, articolo, lotto, scadenza, quantità, UM, layout, famiglia,
 * linea).
 */
#expiringGoodsTable {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed !important;
    min-width: 2000px;
}
#expiringGoodsTable th, #expiringGoodsTable td {
    padding: 4px 3px;
    text-align: center;
    border-bottom: 1px solid #eee;
    white-space: normal !important;
    word-wrap: break-word !important;
    vertical-align: middle;
    box-sizing: border-box;
}
#expiringGoodsTable th {
    background-color: #CFD8DC;
    font-weight: bold;
    color: #333;
    position: sticky;
    top: 0;
    z-index: 5;
}
#expiringGoodsTable input, #expiringGoodsTable select {
    padding: 6px;
    border: 1px solid #cce7f0;
    border-radius: 5px;
    box-sizing: border-box;
    font-size: 0.85em;
    width: 100%;
    text-align: center;
}
/* Larghezze colonne specifiche per Merce in Scadenza */
#expiringGoodsTable th:nth-child(1) { width: 30px; }
#expiringGoodsTable th:nth-child(2) { width: 80px; }
#expiringGoodsTable th:nth-child(3) { width: 430px; }
#expiringGoodsTable th:nth-child(4) { width: 100px; }
#expiringGoodsTable th:nth-child(5) { width: 90px; }
#expiringGoodsTable th:nth-child(6) { width: 90px; }
#expiringGoodsTable th:nth-child(7) { width: 50px; }
#expiringGoodsTable th:nth-child(8) { width: 150px; }
#expiringGoodsTable th:nth-child(9) { width: 150px; }
#expiringGoodsTable th:nth-child(10){ width: 150px; }
/* Allineamento a sinistra per alcune colonne di testo */
#expiringGoodsTable td:nth-child(3) input,
#expiringGoodsTable td:nth-child(9) input,
#expiringGoodsTable td:nth-child(10) input {
    text-align: left;
}
#shippingScheduleTable th:nth-child(14),#arrivalScheduleTable th:nth-child(15){ width: 50px; }
#shippingScheduleTable th:nth-child(15),#arrivalScheduleTable th:nth-child(16){ width: 140px; }


/* Stile per il raggruppamento degli OV nel Gantt Magazzino */
.gantt-ov-group {
    background-color: #0D47A1; /* Blu scuro come richiesto */
    border: 1px solid #1976D2;
    border-radius: 6px;
    padding: 4px;
    margin-bottom: 3px;
    width: 100%;
    box-sizing: border-box;
}
.gantt-ov-group-header {
    color: white;
    font-weight: bold;
    font-size: 0.8em;
    text-align: center;
    margin-bottom: 2px;
    border-bottom: 1px solid #64B5F6;
    padding-bottom: 2px;
}
   

/* Colore giallo intenso per i medical device nel Gantt spedizioni */
.gantt-task.shipping-task.medical-device-shipping {
    background-color: #FFD600; /* Giallo intenso */
    color: #424242; /* Testo più scuro per contrasto */
    border: 1px solid #FFAB00;
}

/* Assicura che il contenitore del tooltip sia visibile */
.generic-tooltip.visible {
    opacity: 1;
    visibility: visible;
}

.control-group-separator {
    border-left: 2px solid #e0e0e0;
    margin: 0 10px;
    align-self: stretch; /* Fa sì che la linea si estenda per tutta l'altezza */
}


/* =================================================================== */
/* ==> BLOCCO CSS PER GANTT MAGAZZINO E LAYOUT (CORRETTO) <== */
/* =================================================================== */

/* 1. Allarga il Gantt e il suo contenitore per occupare più spazio */
#warehouseGanttChartContainer {
    max-width: none;
    width: 100%;
    /* Abilita lo scorrimento orizzontale sul contenitore del Gantt di magazzino.
       In questo modo l'utente può vedere tutte le colonne se la larghezza totale
       supera lo spazio disponibile. */
    /* overflow-x:auto removed by QBAR */
    overflow-y: hidden;
    position: relative; /* per posizionare i pulsanti di scorrimento interni */
}
.warehouse-gantt-chart {
    /* Aggiornato per riflettere la larghezza totale delle 30 colonne (30 × 110px) più l'intestazione.
       Questo impedisce che il grafico venga schiacciato ma consente comunque di visualizzare tutte le
       colonne in un'unica schermata. */
    min-width: 3500px; /* 200px intestazione + 30×110px = 3500px */
}

/* Pulsanti di scorrimento per il Gantt di magazzino.  Questi bottoni sono
   posizionati ai lati del contenitore e permettono all'utente di
   spostarsi orizzontalmente tra le colonne senza utilizzare la barra
   di scorrimento.  Le frecce sono grandi e leggermente trasparenti per
   non disturbare la visualizzazione del grafico. */
.gantt-scroll-btn {
    position: absolute;
    top: 50%;
    transform: translateY(-50%);
    /* Aumentato lo z-index per garantire che i pulsanti restino visibili
       sopra altri elementi come overlay o tooltip. */
    z-index: 150;
    width: 36px;
    height: 36px;
    line-height: 36px;
    text-align: center;
    border: none;
    border-radius: 50%;
    background-color: rgba(255, 255, 255, 0.8);
    box-shadow: 0 2px 6px rgba(0,0,0,0.2);
    cursor: pointer;
    font-size: 24px;
    color: #333;
    padding: 0;
}
#warehouseGanttScrollLeft {
    left: 5px;
}
#warehouseGanttScrollRight {
    right: 5px;
}

/* 2. Stile per i gruppi di SPEDIZIONE (blu) */
.gantt-ov-group.shipping-group {
    background-color: #E3F2FD; /* Sfondo azzurro chiaro */
    border: 2px solid #0D47A1; /* Bordo blu scuro */
}
.gantt-ov-group.shipping-group .gantt-ov-group-header {
    background-color: #0D47A1; /* Intestazione blu scuro */
    color: white;
    border-bottom: 1px solid #64B5F6;
}

/* 3. Stile per i gruppi di ARRIVO (verde) */
.gantt-ov-group.arrival-group {
    background-color: #E8F5E9; /* Sfondo verde chiaro */
    border: 2px solid #1B5E20; /* Bordo verde scuro */
}
.gantt-ov-group.arrival-group .gantt-ov-group-header {
    background-color: #1B5E20; /* Intestazione verde scuro */
    color: white;
    border-bottom: 1px solid #81C784;
}

/* 4. Stile per le icone di layout nel Gantt */
.layout-icon {
    font-size: 0.9em;
    font-weight: bold;
    margin-left: 8px;
    color: #FFEB3B; /* Giallo per alta visibilità */
    display: inline-block;
}
.layout-icon .thermo-red {
    color: #E53935; /* Rosso per il termometro */
}
.layout-icon .snow-blue {
    color: #81D4FA; /* Azzurro per la neve */
}

/* Colore giallo più scuro per i medical device specifici in spedizione */
.gantt-task.shipping-task.medical-device-shipping-priority {
    background-color: #FFC107; /* Giallo più scuro/ambra */
    color: #000000; /* Testo nero per leggibilità */
    font-weight: bold;
    border: 2px solid #FF8F00;
}

/* Stile per l'icona di priorità (puntino rosso lampeggiante) */
.priority-icon {
    width: 10px;
    height: 10px;
    background-color: red;
    border-radius: 50%;
    display: inline-block;
    margin-right: 5px;
    animation: blink-animation 1s infinite;
    vertical-align: middle;
}

@keyframes blink-animation {
    0% { opacity: 1; }
    50% { opacity: 0.2; }
    100% { opacity: 1; }
}

/* Stile per il flag di avvenuto caricamento file */
.file-status-flag {
    position: absolute;
    top: -8px;
    right: -8px;
    background-color: #4CAF50; /* Verde brillante */
    color: white;
    border-radius: 50%;
    width: 20px;
    height: 20px;
    display: flex; /* Utilizzato per centrare perfettamente la spunta */
    align-items: center;
    justify-content: center;
    font-size: 14px;
    line-height: 1;
    font-weight: bold;
    box-shadow: 0 1px 3px rgba(0,0,0,0.2);
    border: 1px solid white;
    cursor: help; /* Indica che si possono avere più info al passaggio del mouse */
    display: none; /* Nascosto di default, viene mostrato da JavaScript */
}

/* Piccola etichetta che mostra l'ultima data/ora di importazione dei file. */
.last-import-time {
    font-size: 0.75em;
    /* Per evitare che il testo dell'ultimo import sposti i pulsanti, 
       riserviamo sempre una larghezza minima e non usiamo margini laterali.
       In questo modo il posto rimane, anche se la data non è presente. */
    display: inline-block;
    min-width: 130px;
    color: #333;
    vertical-align: middle;
}

/*
 * Contenitore riassuntivo per mostrare le date/ore degli ultimi import.
 * Questa sezione non sposta altri elementi perché ha uno spazio dedicato.
 */
.last-import-summary {
    display: block;
    min-height: 20px;
}
.last-import-summary div {
    margin-bottom: 2px;
}

/* =================================================================== */
/* ==> BLOCCO CSS PER GANTT MAGAZZINO E LEGGENDA FAMIGLIE (AGGIORNATO) <== */
/* =================================================================== */

/* 1. Stili per i NUOVI colori delle famiglie (più distinti) */
/* Lo sfondo viene applicato direttamente al .gantt-task */
.gantt-task.family-color-0 { background-color: #FFCDD2; color: #B71C1C; border: 1px solid #B71C1C; } /* Rosso Chiaro */
.gantt-task.family-color-1 { background-color: #C5CAE9; color: #1A237E; border: 1px solid #1A237E; } /* Indaco Chiaro */
.gantt-task.family-color-2 { background-color: #B2DFDB; color: #004D40; border: 1px solid #004D40; } /* Teal Chiaro */
.gantt-task.family-color-3 { background-color: #FFF9C4; color: #F57F17; border: 1px solid #F57F17; } /* Giallo Chiaro */
.gantt-task.family-color-4 { background-color: #D1C4E9; color: #311B92; border: 1px solid #311B92; } /* Viola Chiaro */
.gantt-task.family-color-5 { background-color: #F8BBD0; color: #880E4F; border: 1px solid #880E4F; } /* Rosa Chiaro */
.gantt-task.family-color-6 { background-color: #D7CCC8; color: #3E2723; border: 1px solid #3E2723; } /* Marrone Chiaro */
.gantt-task.family-color-7 { background-color: #FFE0B2; color: #E65100; border: 1px solid #E65100; } /* Arancione Chiaro */
.gantt-task.family-color-default { background-color: #CFD8DC; color: #263238; border: 1px solid #263238; } /* Grigio di default */

/* Definizione delle larghezze per le nuove colonne OPE e OV nella tabella di produzione.
   Queste colonne vengono inserite all'inizio della tabella Dettaglio Produzione per
   visualizzare rispettivamente l'Ordine di Produzione Esterno (OPE) e l'Ordine di Vendita (OV).
   Impostiamo una larghezza fissa per evitare che le celle si restringano eccessivamente. */
.col-ope,
.col-ov {
    width: 60px;
}

/* Stile per le icone di ordinamento visualizzate nelle intestazioni di colonna.
   Le icone sono piccole frecce che consentono di ordinare rapidamente ogni colonna
   in ordine crescente o decrescente. */
.sort-icon {
    /* Migliora la visibilità delle icone di ordinamento aumentando la dimensione
       e il contrasto. */
    font-size: 0.9em;
    color: #333;
    cursor: pointer;
    margin-left: 4px;
    user-select: none;
}

/* 2. Stile per il gruppo OV (ora trasparente per far risaltare i task colorati) */
.gantt-ov-group.arrival-group {
    background-color: transparent;
    border: none;
    padding: 1px 0; /* Riduciamo lo spazio per compattare */
}

/* 3. Intestazione del gruppo OV (rimane scura per leggibilità) */
.gantt-ov-group.arrival-group .gantt-ov-group-header {
    background-color: #455A64; /* Grigio-blu scuro */
    color: white;
    border-bottom: 1px solid #90A4AE;
}

/* 4. Stili per la legenda nel Gantt (invariati ma funzioneranno con i nuovi colori) */
.gantt-legend {
    margin-top: 8px;
    padding-left: 5px;
    font-size: 0.8em;
    text-align: left;
}
.legend-item {
    display: flex;
    align-items: center;
    margin-bottom: 3px;
}
.legend-color-box {
    width: 15px;
    height: 15px;
    margin-right: 6px;
    border: 1px solid #777;
    flex-shrink: 0;
}
.legend-text {
    font-weight: normal;
}

/* Stile per il bordo nero speciale richiesto */
.gantt-task.gantt-task-special-border {
    border: 3px solid black !important;
}

/* === STILI SPECIFICI PER LA STAMPA ISOLATA DEL PROGRAMMA GIORNALIERO === */
@media print {
    /* Regola #1: Nasconde tutti gli elementi principali quando stampiamo il programma giornaliero */
    body.printing-daily-production .container > *:not(#dailyProductionContainer) {
        display: none !important;
    }

    /* Regola #2: Assicura che il contenitore del programma giornaliero sia visibile */
    body.printing-daily-production #dailyProductionContainer {
        display: block !important;
        box-shadow: none !important;
        border: none !important;
        padding: 0 !important;
        margin: 0 !important;
    }

    /* Regola #3: Nasconde i pulsanti e i filtri all'interno della sezione durante la stampa */
    body.printing-daily-production #dailyProductionContainer .daily-production-controls {
        display: none !important;
    }
    
    /* Regola #4: Nasconde l'header e la barra dei comandi principali */
    body.printing-daily-production header,
    body.printing-daily-production .sticky-controls-wrapper {
        display: none !important;
    }

    /* Regola #5: Traccia una riga nera tra le righe della tabella giornaliera
       per facilitare la compilazione manuale dell'esito.  Applichiamo il
       bordo inferiore alle celle, perché la tabella utilizza border-collapse
       che annulla i bordi impostati sulle righe.  Inoltre, per tabelle
       esportate come PDF (programma giornaliero, programmazione OPI, ecc.),
       aggiungiamo il bordo alle celle di qualsiasi tabella all'interno del
       contenitore di stampa. */
    body.printing-daily-production #dailyProductionContainer table.daily-production-table tbody td {
        border-bottom: 1px solid #000 !important;
    }
    /* Applica il bordo inferiore a tutte le celle nelle tabelle quando si
       stampa il programma giornaliero o altri programmi di produzione.  Questo
       copre eventuali tabelle aggiuntive create dinamicamente per la stampa
       (es. Programma di produzione con colonne OP/OV) che potrebbero non
       avere la classe .daily-production-table. */
    body.printing-daily-production #dailyProductionContainer table tbody td {
        border-bottom: 1px solid #000 !important;
    }
}

/* =================================================================== */
/* ==> NUOVI STILI PER COMMENTI QA NELLE SPEDIZIONI <== */
/* =================================================================== */

/* Stile per il contenitore principale del tooltip che ora può avere 3 box */
.tooltip-container {
    display: flex;
    gap: 8px;
    padding: 5px;
    align-items: stretch; /* Assicura che i box abbiano la stessa altezza */
}

/* Stile per il nuovo box dei commenti QA */
.qa-comments-tooltip {
    background-color: #FFFDE7; /* Giallo molto chiaro */
    border-color: #FFD54F;
    color: #4E342E;
    display: flex;
    flex-direction: column; /* Organizza contenuto verticalmente */
}
.qa-comments-tooltip h3 {
    border-bottom-color: #FFF176;
    color: #BF360C;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.qa-comments-tooltip p {
    flex-grow: 1; /* Fa in modo che il paragrafo occupi lo spazio disponibile */
    white-space: pre-wrap; /* Mantiene la formattazione del testo (a capo, etc.) */
    text-align: left;
}

/* Stile per le icone del lucchetto */
.lock-icon {
    cursor: pointer;
    font-size: 1.2em;
    margin-left: 10px;
}

/* Stile per il modale di inserimento password e commenti */
.qa-modal-content {
    background-color: #fff;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
    text-align: center;
    max-width: 450px;
    width: 90%;
}
.qa-modal-content h3 {
    color: #BF360C;
    margin-top: 0;
}
.qa-modal-content p {
    color: #555;
    margin-bottom: 15px;
}
.qa-modal-content input,
.qa-modal-content textarea {
    width: calc(100% - 24px);
    padding: 12px;
    margin-bottom: 15px;
    border: 1px solid #ddd;
    border-radius: 8px;
    font-size: 1em;
    font-family: 'Quicksand', sans-serif;
}
.qa-modal-content textarea {
    min-height: 120px;
    resize: vertical;
    text-align: left;
}
.qa-modal-buttons {
    display: flex;
    justify-content: center;
    gap: 15px;
}

/* Stile per rendere il contenitore del task un punto di riferimento */
.gantt-task {
    position: relative; /* FONDAMENTALE per posizionare il lucchetto all'interno */
}

/* Stile per il nuovo lucchetto posizionato FUORI dal task del Gantt */
.gantt-qa-lock {
    position: absolute;
    top: -8px;           /* Sposta l'icona 8px SOPRA la barra */
    right: -8px;         /* Sposta l'icona 8px a DESTRA della barra */
    font-size: 1.2em;
    cursor: pointer;
    z-index: 10;
    padding: 3px;
    background-color: #ffffff; /* Sfondo bianco per farlo risaltare */
    border: 1px solid #b0bec5; /* Bordo grigio chiaro */
    border-radius: 50%;      /* Lo rende rotondo */
    box-shadow: 0 1px 4px rgba(0,0,0,0.25); /* Ombra per dare profondità */
    line-height: 1;
    transition: transform 0.2s ease;
}

.gantt-qa-lock:hover {
    transform: scale(1.25); /* Ingrandisce di più al passaggio del mouse */
}

/* ========================================================================= */
/* ==> STILI PER LO STATO CQ (pallini colorati a sinistra dei task) <== */
/* ========================================================================= */
/* Il pallino CQ viene posizionato in alto a sinistra del task e mostra lo   */
/* stato dell'accettazione qualità: bianco = da analizzare, giallo = accettato*/
/* con deroga, verde = conforme, rosso = non conforme. Solo l'operatore CQ   */
/* autorizzato può modificarlo inserendo la password.                         */
.cq-status-dot {
    position: absolute;
    top: -8px;
    left: -8px;
    width: 14px;
    height: 14px;
    border-radius: 50%;
    border: 1px solid #b0bec5;
    box-shadow: 0 1px 4px rgba(0,0,0,0.25);
    cursor: pointer;
    z-index: 10;
}
/* Colori per gli stati CQ */
.cq-status-white { background-color: #ffffff; }
.cq-status-yellow {
    /* Colore giallo più intenso per migliorare il contrasto con lo sfondo */
    background-color: #FFC107;
}

/* ========================================================================= */
/* ==> STILI PER LO STATO MAGAZZINO (pallini per merce in arrivo) <== */
/* ========================================================================= */
/* Il pallino Magazzino viene posizionato in alto a sinistra del task per la
   merce in arrivo e permette di passare dallo stato "da evadere" (bianco)
   allo stato "evasa" (verde). Solo gli utenti con password corretta possono
   modificarlo. */
.mag-status-dot {
    position: absolute;
    top: -8px;
    left: -8px;
    width: 14px;
    height: 14px;
    border-radius: 50%;
    border: 1px solid #b0bec5;
    box-shadow: 0 1px 4px rgba(0,0,0,0.25);
    cursor: pointer;
    z-index: 10;
}
/* Colori per lo stato Magazzino */
.mag-status-white { background-color: #ffffff; }
.mag-status-green { background-color: #4CAF50; }

/* Stile della legenda Magazzino */
.mag-legend {
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 4px;
    font-size: 0.85em;
    margin-top: 8px;
    margin-right: 0;
}
.mag-legend-item {
    display: flex;
    align-items: center;
    gap: 5px;
}
.mag-legend .mag-status-dot {
    position: static;
    display: inline-block;
    margin-right: 4px;
}
.mag-legend-title {
    font-weight: bold;
}

/* ========================================================================= */
/* ==> STILI PER L'AVVISO MAGAZZINO <== */
/* ========================================================================= */
#warehouseNotification {
    display: none;
    position: fixed;
    top: 25%;
    left: 50%;
    transform: translateX(-50%);
    background: #ffffff;
    border: 2px solid #388E3C; /* verde per magazzino */
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
    padding: 20px;
    z-index: 10000;
    max-width: 500px;
    font-size: 0.9em;
    border-radius: 4px;
    pointer-events: auto;
    display: flex;
    flex-direction: column;
    max-height: 70vh;
    cursor: move;
}
#warehouseNotification p {
    margin-bottom: 15px;
    color: #388E3C;
    font-weight: bold;
}
#warehouseNotification .warehouse-alert-content {
    flex: 1 1 auto;
    overflow-y: auto;
    margin-bottom: 10px;
}
#warehouseNotification .warehouse-alert-buttons {
    display: flex;
    justify-content: flex-end;
    gap: 10px;
    margin-top: 10px;
}
#warehouseNotification .warehouse-alert-buttons button {
    padding: 5px 12px;
    border: none;
    border-radius: 3px;
    cursor: pointer;
}
#warehouseNotification .warehouse-alert-buttons button:first-child {
    background: #f0f0f0;
    color: #333;
}
#warehouseNotification .warehouse-alert-buttons button:last-child {
    background: #388E3C;
    color: #fff;
}
#warehouseNotification .warehouse-close-btn {
    position: absolute;
    top: 4px;
    right: 6px;
    cursor: pointer;
    font-size: 18px;
    line-height: 18px;
    color: #388E3C;
}
#warehouseNotification ul {
    margin: 0 0 10px 0;
    padding-left: 20px;
}
#warehouseNotification li {
    margin-bottom: 4px;
}
.cq-status-green { background-color: #4CAF50; }
.cq-status-red { background-color: #F44336; }
/* Stile della legenda CQ */
.cq-legend {
    /* Mostra gli elementi della legenda CQ uno sotto l'altro per una maggiore chiarezza */
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 4px;
    font-size: 0.85em;
    margin-top: 8px;
    /* Non utilizzare margin-right; il separatore e margin-top della legenda QA gestiscono lo spacing */
    margin-right: 0;
}
.cq-legend-item {
    display: flex;
    align-items: center;
    gap: 5px;
}

/* I pallini all'interno della legenda devono essere statici e visibili in linea
   con il testo, quindi sovrascriviamo il posizionamento assoluto usato nei task */
.cq-legend .cq-status-dot {
    position: static;
    display: inline-block;
    margin-right: 4px;
}
/* Stile per il titolo della legenda CQ */
.cq-legend-title {
    font-weight: bold;
    margin-right: 8px;
}

        /* ============================================================ */
        /* ==  Nuovi stili per lo stato QA (Quality Assurance)        == */
        /* ============================================================ */
        /* Bandierina QA posizionata nell'angolo in basso a sinistra del task.
           Mostra la lettera "V" per indicare la verifica e cambia colore
           in base allo stato (bianco = in valutazione, giallo = con deroga,
           verde = conforme, rosso = non conforme). */
        .qa-status-flag {
            width: 14px;
            height: 14px;
            border-radius: 2px;
            border: 1px solid #b0bec5;
            box-shadow: 0 1px 4px rgba(0,0,0,0.25);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 10px;
            line-height: 1;
            cursor: pointer;
            position: absolute;
            bottom: -8px;
            /* Posiziona la bandierina QA nell'angolo in basso a sinistra del task */
            left: -8px;
            right: auto;
            top: auto;
            z-index: 10;
        }
        /* Colori per gli stati QA */
        /* Stato "bianco" per QA: torna a essere bianco come in origine, con bordo per
           evidenziarlo sul Gantt. Il colore del testo è ininfluente perché non c'è testo. */
        .qa-status-white { background-color: #ffffff; color: #000; }
        .qa-status-yellow { background-color: #FFC107; color: #000; }
        .qa-status-green { background-color: #4CAF50; color: #fff; }
        .qa-status-red { background-color: #F44336; color: #fff; }

        /* ============================================================ */
        /* ==  Nuovi stili per lo stato di spedizione (ship-status)   == */
        /* ============================================================ */
        /* La bandierina di spedizione viene posizionata in basso a destra
           del task per indicare se l'ordine è stato spedito.  La lettera
           "S" viene mostrata all'interno per chiarezza.  Lo stato
           "white" rappresenta "non spedito", mentre lo stato "green"
           rappresenta "spedito". */
        .ship-status-flag {
            width: 14px;
            height: 14px;
            border-radius: 2px;
            border: 1px solid #b0bec5;
            box-shadow: 0 1px 4px rgba(0,0,0,0.25);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 10px;
            line-height: 1;
            cursor: pointer;
            position: absolute;
            bottom: -8px;
            left: auto;
            right: -8px;
            top: auto;
            z-index: 10;
        }
        /* Colori per gli stati di spedizione */
        .ship-status-white { background-color: #ffffff; color: #000; }
        .ship-status-green { background-color: #2196F3; color: #fff; }

        /* ============================================================ */
        /* ==  Stili per l'avviso di spedizione (shippingNotification) == */
        /* ============================================================ */
        #shippingNotification {
            display: none;
            position: fixed;
            top: 30%;
            left: 50%;
            transform: translateX(-50%);
            background: #ffffff;
            border: 2px solid #2196F3;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
            padding: 20px;
            z-index: 10000;
            max-width: 500px;
            font-size: 0.9em;
            border-radius: 4px;
            pointer-events: auto;
            display: flex;
            flex-direction: column;
            max-height: 70vh;
            cursor: move;
        }
        #shippingNotification p {
            margin-bottom: 15px;
            color: #0d47a1;
            font-weight: bold;
        }
        #shippingNotification .shipping-alert-content {
            flex: 1 1 auto;
            overflow-y: auto;
            margin-bottom: 10px;
        }
        #shippingNotification .shipping-alert-buttons {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 10px;
        }
        #shippingNotification .shipping-alert-buttons button {
            padding: 5px 12px;
            border: none;
            border-radius: 3px;
            cursor: pointer;
        }
        #shippingNotification .shipping-alert-buttons button:first-child {
            background: #f0f0f0;
            color: #333;
        }
        #shippingNotification .shipping-alert-buttons button:last-child {
            background: #2196F3;
            color: #fff;
        }
        #shippingNotification .shipping-close-btn {
            position: absolute;
            top: 4px;
            right: 6px;
            cursor: pointer;
            font-size: 18px;
            line-height: 18px;
            color: #2196F3;
        }
        #shippingNotification ul {
            margin: 0 0 10px 0;
            padding-left: 20px;
        }

        /* Indicatore ADR lampeggiante da visualizzare sui task di spedizione che
           richiedono il trasporto di merci pericolose.  È posizionato in alto a destra
           del task e lampeggia per attirare l'attenzione. */
        .adr-indicator {
            position: absolute;
            top: -8px;
            right: -6px;
            background-color: #F44336;
            color: #fff;
            padding: 0 3px;
            font-size: 8px;
            font-weight: bold;
            border-radius: 2px;
            animation: adr-blink 1s steps(2, start) infinite;
            z-index: 20;
        }
        @keyframes adr-blink {
            0% { opacity: 1; }
            50% { opacity: 0; }
            100% { opacity: 1; }
        }

        /* Evidenzia le spedizioni ADR con una cornice rossa lampeggiante.  La
           classe viene applicata al task di spedizione quando il codice
           articolo appartiene all'elenco ADR.  La cornice lampeggia per
           attirare l'attenzione e non interferisce con il contenuto interno. */
        .adr-shipping {
            border: 2px solid #F44336 !important;
            animation: adr-border-blink 1s linear infinite;
        }
        @keyframes adr-border-blink {
            0%, 100% {
                box-shadow: 0 0 0 2px rgba(244,67,54,1);
            }
            50% {
                box-shadow: 0 0 0 2px rgba(244,67,54,0);
            }
        }

        /* Stile per l'avviso ADR che appare come pop-up.  Il pop-up è
           centrato e contiene due pulsanti: posponi e ok. */
        #adrNotification {
            /* L'avviso ADR è nascosto di default; sarà reso visibile via JS
               quando vengono rilevate spedizioni ADR. */
            display: none;
            position: fixed;
            top: 20%;
            left: 50%;
            transform: translateX(-50%);
            background: #ffffff;
            border: 2px solid #F44336;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
            padding: 20px;
            z-index: 10000;
            max-width: 400px;
            font-size: 0.9em;
            border-radius: 4px;
            /* Consente l'interazione con i pulsanti all'interno del pop-up */
            pointer-events: auto;
            /* Layout a colonne: il contenuto scorrevole e i pulsanti
               sono separati verticalmente.  La finestra si adatta allo
               schermo e i pulsanti restano sempre visibili. */
            flex-direction: column;
            max-height: 70vh;
        }
        #adrNotification p {
            margin-bottom: 15px;
            color: #F44336;
            font-weight: bold;
        }
        /* Contenuto scrollabile del pop-up ADR: include titolo e lista.
           Grazie a flex:1 e overflow-y:auto, questo elemento occupa
           tutto lo spazio disponibile lasciando i pulsanti sempre
           visibili nella parte bassa della finestra. */
        #adrNotification .adr-alert-content {
            flex: 1 1 auto;
            overflow-y: auto;
            margin-bottom: 10px;
        }
        /* Pulsanti dell'avviso ADR: sono allineati a destra e
           separati da uno spazio.  Restano sempre visibili anche con
           contenuti lunghi. */
        #adrNotification .adr-alert-buttons {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 10px;
        }
        #adrNotification .adr-alert-buttons button {
            padding: 5px 12px;
            border: none;
            border-radius: 3px;
            cursor: pointer;
        }
        #adrNotification .adr-alert-buttons button:first-child {
            background: #f0f0f0;
            color: #333;
        }
        #adrNotification .adr-alert-buttons button:last-child {
            background: #F44336;
            color: #fff;
        }

        /* Pulsante di chiusura per l'avviso ADR: posizionato nell'angolo in
           alto a destra della finestra.  Permette all'utente di
           chiudere l'avviso immediatamente, indipendentemente dai
           pulsanti Posponi/OK. */
        #adrNotification .adr-close-btn {
            position: absolute;
            top: 4px;
            right: 6px;
            cursor: pointer;
            font-size: 18px;
            line-height: 18px;
            color: #F44336;
        }

        /* Override e perfeziona lo stile dell'avviso ADR.  Definiamo un
           bordo e dimensioni maggiorate e aggiungiamo stili per gli
           elementi della lista (ul/li) inclusi nel pop-up. */
        #adrNotification {
            /* Aumenta leggermente la larghezza massima del pop‑up ADR.
               Non ridefinire overflow‑y qui: lo scroll verticale è
               gestito dal contenitore .adr-alert-content per mantenere
               sempre visibili i pulsanti. */
            max-width: 500px;
            max-height: 70vh;
        }
        #adrNotification ul {
            margin: 0 0 10px 0;
            padding-left: 20px;
        }
        #adrNotification li {
            margin-bottom: 4px;
        }

        /* == STILI PER L'AVVISO CQ/QA == */
        #qualityNotification {
            display: none;
            position: fixed;
            top: 25%;
            left: 50%;
            transform: translateX(-50%);
            background: #ffffff;
            border: 2px solid #1976D2; /* Blu medio per CQ/QA */
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
            padding: 20px;
            z-index: 10000;
            max-width: 500px;
            font-size: 0.9em;
            border-radius: 4px;
            pointer-events: auto;
            display: flex;
            flex-direction: column;
            max-height: 70vh;
        }
        #qualityNotification p {
            margin-bottom: 15px;
            color: #1976D2;
            font-weight: bold;
        }
        #qualityNotification .quality-alert-content {
            flex: 1 1 auto;
            overflow-y: auto;
            margin-bottom: 10px;
        }
        #qualityNotification .quality-alert-buttons {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 10px;
        }
        #qualityNotification .quality-alert-buttons button {
            padding: 5px 12px;
            border: none;
            border-radius: 3px;
            cursor: pointer;
        }
        #qualityNotification .quality-alert-buttons button:first-child {
            background: #f0f0f0;
            color: #333;
        }
        #qualityNotification .quality-alert-buttons button:last-child {
            background: #1976D2;
            color: #fff;
        }
        #qualityNotification .quality-close-btn {
            position: absolute;
            top: 4px;
            right: 6px;
            cursor: pointer;
            font-size: 18px;
            line-height: 18px;
            color: #1976D2;
        }
        #qualityNotification ul {
            margin: 0 0 10px 0;
            padding-left: 20px;
        }
        #qualityNotification li {
            margin-bottom: 4px;
        }
        /* Rende l'intero pop-up trascinabile mostrando un cursore di spostamento.  Gli
           elementi interattivi (pulsanti e icone di chiusura) manterranno il
           loro cursore predefinito grazie al codice JS che ignora il drag. */
        #adrNotification,
        #qualityNotification {
            cursor: move;
        }
        #adrNotification button,
        #qualityNotification button,
        #adrNotification .adr-close-btn,
        #qualityNotification .quality-close-btn {
            cursor: pointer;
        }
        /* Stile della legenda QA */
.qa-legend {
        /* Dispone gli elementi della legenda QA in colonna. Viene separata dalla legenda CQ
           da un separatore HTML e da un margine superiore più ampio */
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        gap: 4px;
        font-size: 0.85em;
        /* Aggiungiamo un margine superiore significativo per distanziare la legenda QA dalla CQ */
        margin-top: 16px;
        /* Non applichiamo spaziatura orizzontale; le legende sono impilate verticalmente */
        margin-left: 0;
        }
        .qa-legend-item {
            display: flex;
            align-items: center;
            gap: 5px;
        }
        /* Le bandierine all'interno della legenda devono essere statiche e visibili in linea */
        .qa-legend .qa-status-flag {
            position: static;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            margin-right: 4px;
        }
        .qa-legend-title {
            font-weight: bold;
            margin-right: 8px;
        }

        /* Stile per il riquadro delle note interne (importazione colonna Q nei task di arrivo).
           Utilizza colori pastello per il bordo e lo sfondo per renderlo evidente ma non invasivo. */
        .arrival-note-tooltip {
            border-color: #FFE082;
            background-color: #FFF8E1;
        }
        /* Il titolo "Nota interna" all'interno del riquadro deve avere un colore scuro
           per contrastare con lo sfondo pastello. */
        .arrival-note-tooltip h3 {
            color: #333;
        }
        /* Piccolo quadratino colorato utilizzato accanto alla dicitura "Nota interna" */
        .internal-note-icon {
            display: inline-block;
            width: 10px;
            height: 10px;
            background-color: #FFE082;
            margin-right: 6px;
            border-radius: 2px;
        }

        /* Blocco inline per le note interne inserite nel Dettaglio Ordine.  Usa colori pastello
           coerenti con quelli della tooltip delle note di arrivo e un bordo laterale per
           evidenziare la sezione.  Viene posizionato sotto le informazioni principali del
           dettaglio ordine. */
        .internal-note-inline {
            border-left: 4px solid #FFE082;
            background-color: #FFF8E1;
            padding: 6px;
            margin-top: 6px;
            border-radius: 4px;
        }
        .internal-note-inline h4 {
            margin: 0 0 4px 0;
            font-size: 14px;
            color: #333;
            display: flex;
            align-items: center;
        }
        .internal-note-inline p {
            margin: 0;
            color: #333;
        }


/* =================================================================== */
/* ==> STILI PER LA STAMPA DEL GANTT MAGAZZINO <== */
/* =================================================================== */

@media print {
    /* Regola #1: Nasconde TUTTO di default quando la classe di stampa è attiva */
    body.printing-warehouse-gantt > .container > *:not(#warehouseGanttChartContainer) {
        display: none !important;
    }
    body.printing-warehouse-gantt header,
    body.printing-warehouse-gantt .sticky-controls-wrapper {
        display: none !important;
    }

    /* Regola #2: Mostra SOLO il contenitore del Gantt Magazzino */
    body.printing-warehouse-gantt #warehouseGanttChartContainer {
        display: block !important;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        box-shadow: none !important;
        border: none !important;
        padding: 0 !important;
        margin: 0 !important;
    }

    /* Regola #3: Nasconde gli elementi interni non necessari (titolo e pulsante) */
    body.printing-warehouse-gantt #warehouseGanttChartContainer > div:first-child {
        display: none !important;
    }

    /* Regola #4: Formatta il grafico per riempire la pagina */
    body.printing-warehouse-gantt .warehouse-gantt-chart {
        min-width: unset !important;
        font-size: 7pt; /* Riduci la dimensione del testo per far entrare più informazioni */
    }
    
    body.printing-warehouse-gantt .gantt-task {
        line-height: 1.1;
        padding: 1px;
        font-size: 0.8em;
    }

    /* Regola #5: Imposta il layout della pagina di stampa */
    /* L'istruzione @page non può essere annidata in un selettore;
       la definiamo direttamente a livello di @media print per evitare errori CSS. */
    @page {
        size: A4 landscape;
        margin: 1cm;
    }
}


/* =================================================================== */
/* ==> REGOLE SPECIFICHE PER ALLINEARE LE COLONNE DI "MERCE NON ARRIVATA" <== */
/* =================================================================== */

/* Imposta la larghezza minima totale uguale a quella della tabella Arrivi */
#overdueArrivalsTable {
    min-width: 2450px;
}

/* Applica le stesse larghezze delle colonne della tabella Arrivi */
#overdueArrivalsTable th:nth-child(1) { width: 30px; }
#overdueArrivalsTable th:nth-child(2) { width: 80px; }
#overdueArrivalsTable th:nth-child(3) { width: 100px; }
#overdueArrivalsTable th:nth-child(4) { width: 430px; }
#overdueArrivalsTable th:nth-child(5) { width: 150px; } /* Layout */
#overdueArrivalsTable th:nth-child(6) { width: 90px; }
#overdueArrivalsTable th:nth-child(7) { width: 50px; }
#overdueArrivalsTable th:nth-child(8) { width: 90px; }
#overdueArrivalsTable th:nth-child(9) { width: 90px; }
#overdueArrivalsTable th:nth-child(10){ width: 300px; }
#overdueArrivalsTable th:nth-child(11){ width: 200px; }
#overdueArrivalsTable th:nth-child(12){ width: 320px; }
#overdueArrivalsTable th:nth-child(13){ width: 70px; }
#overdueArrivalsTable th:nth-child(14){ width: 140px; }
#overdueArrivalsTable th:nth-child(15){ width: 50px; }

/* =================================================================== */
/* ==> REGOLE SPECIFICHE PER ALLINEARE LE COLONNE DI "MERCE IN QUARANTENA" <== */
/* Copia le larghezze della tabella "Merce non Arrivata" così che la
   tabella quarantena abbia esattamente la stessa larghezza e allineamento. */
#quarantineTable {
    /* Imposta la larghezza minima totale uguale a quella delle altre tabelle */
    min-width: 2450px;
}
#quarantineTable th:nth-child(1) { width: 30px; }
#quarantineTable th:nth-child(2) { width: 80px; }
#quarantineTable th:nth-child(3) { width: 100px; }
#quarantineTable th:nth-child(4) { width: 430px; }
#quarantineTable th:nth-child(5) { width: 150px; }
#quarantineTable th:nth-child(6) { width: 90px; }
#quarantineTable th:nth-child(7) { width: 50px; }
#quarantineTable th:nth-child(8) { width: 90px; }
#quarantineTable th:nth-child(9) { width: 90px; }
#quarantineTable th:nth-child(10){ width: 300px; }
#quarantineTable th:nth-child(11){ width: 200px; }
#quarantineTable th:nth-child(12){ width: 320px; }
#quarantineTable th:nth-child(13){ width: 70px; }
#quarantineTable th:nth-child(14){ width: 140px; }
#quarantineTable th:nth-child(15){ width: 50px; }
#overdueArrivalsTable th:nth-child(16){ width: 140px; }


/* Stile larghezza colonne per tabella OPI */
#opiTable th, #opiTable td {
    text-align: center;
    vertical-align: middle;
    padding: 6px 3px;
    white-space: normal;
    word-break: break-word;
}

#opiTable th:nth-child(1),#opiTable td:nth-child(1) { width: 30px; } /* Operatore */
#opiTable th:nth-child(2), #opiTable td:nth-child(2) { width: 30px;  }   /* Vuoto/checkbox */
#opiTable th:nth-child(3), #opiTable td:nth-child(3) { width: 30px;  }   /* OV */
#opiTable th:nth-child(4), #opiTable td:nth-child(4) { width: 30px;  }   /* OP */
#opiTable th:nth-child(5), #opiTable td:nth-child(5) { width: 150px;  }   /* Codice */
#opiTable th:nth-child(6), #opiTable td:nth-child(6) { width: 150px; }   /* Articolo */
#opiTable th:nth-child(7), #opiTable td:nth-child(7) { width: 30px; }   /* Cliente */
#opiTable th:nth-child(8), #opiTable td:nth-child(8) { width: 30px;  }   /* Lotto */
#opiTable th:nth-child(9), #opiTable td:nth-child(9) { width: 20px;  }   /* Quantità */
#opiTable th:nth-child(10), #opiTable td:nth-child(10) { width: 60px;  }   /* UM */
#opiTable th:nth-child(11), #opiTable td:nth-child(11) { width: 30px;  } /* Data Produzione */
#opiTable th:nth-child(12), #opiTable td:nth-child(12) { width: 50px; } /* Scadenza Lotto */
#opiTable th:nth-child(13), #opiTable td:nth-child(13) { width: 30px;  } /* Log Mov. */

 
/* === QBAR OVERRIDES: place a single native scrollbar between the Gantt and Quarantena === */
#warehouseGanttChartContainer { overflow-x: hidden !important; overflow-y: visible; }
#warehouseGanttScrollWrapper {
  width: 100%;
  overflow-x: auto;
  /* Abilita lo scorrimento verticale per mantenere l'intestazione dei giorni
     sempre visibile e limita l'altezza della griglia.  Quando il contenuto
     supera l'altezza del contenitore, compare una barra di scorrimento
     verticale.  L'altezza è impostata al 60% della viewport per adattarsi
     dinamicamente alla finestra. */
  overflow-y: auto;
  max-height: 60vh;
  margin: 8px 0 14px 0; /* sits clearly between Gantt grid and Quarantena title */
  scrollbar-gutter: stable both-edges;
}
#warehouseGanttScrollWrapper > .gantt-chart.warehouse-gantt-chart {
  min-width: 3500px; /* 30 cols × 110px + 200px header (already set elsewhere; this is a safe guard) */
}

/* === Navigazione rapida verticale (QuickNav) === */
.quick-nav-vertical {
  position: fixed;
  top: 50%;
  right: 8px;
  transform: translateY(-50%);
  display: flex;
  flex-direction: column;
  align-items: center;
  z-index: 1000;
}
.quick-nav-vertical a {
  display: block;
  width: 28px;
  height: 28px;
  margin: 4px 0;
  border-radius: 50%;
  background-color: #37474F;
  color: #fff;
  text-align: center;
  line-height: 28px;
  font-size: 14px;
  text-decoration: none;
  position: relative;
}
.quick-nav-vertical a:hover {
  background-color: #455A64;
}
.quick-nav-vertical a .tooltip {
  visibility: hidden;
  opacity: 0;
  position: absolute;
  right: 36px;
  top: 50%;
  transform: translateY(-50%);
  background: rgba(0, 0, 0, 0.75);
  color: #fff;
  padding: 4px 8px;
  border-radius: 4px;
  white-space: nowrap;
  transition: opacity 0.2s;
  font-size: 12px;
  pointer-events: none;
}
.quick-nav-vertical a:hover .tooltip {
  visibility: visible;
  opacity: 1;
}

/* Stile dedicato al pulsante di logout nella navigazione verticale.  Questo
   pulsante riutilizza le dimensioni delle altre icone ma con un colore
   distintivo per indicare l'azione di uscita.  Si presenta come un
   cerchio rosso con un'icona di accensione bianca. */
.logout-nav-btn {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 28px;
  height: 28px;
  margin: 4px 0;
  border-radius: 50%;
  background-color: #D32F2F;
  color: #fff;
  text-decoration: none;
  font-size: 16px;
  position: relative;
}
.logout-nav-btn:hover {
  background-color: #B71C1C;
}
/* Reimpiega lo stile del tooltip per il logout: posizioniamo a sinistra
   come per gli altri elementi della quick nav */
.logout-nav-btn .tooltip {
  visibility: hidden;
  opacity: 0;
  position: absolute;
  right: 36px;
  top: 50%;
  transform: translateY(-50%);
  background: rgba(0, 0, 0, 0.75);
  color: #fff;
  padding: 4px 8px;
  border-radius: 4px;
  white-space: nowrap;
  transition: opacity 0.2s;
  font-size: 12px;
  pointer-events: none;
}
.logout-nav-btn:hover .tooltip {
  visibility: visible;
  opacity: 1;
}

/* === Modal Sblocco CQ/QA === */
.sblocco-modal {
  display: none;
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: rgba(0,0,0,0.6);
  z-index: 2000;
  justify-content: center;
  align-items: flex-start;
  overflow-y: auto;
  padding-top: 40px;
}
.sblocco-content {
  background: #fff;
  border-radius: 6px;
  width: 95%;
  max-width: 1300px;
  padding: 20px;
  box-shadow: 0 4px 10px rgba(0,0,0,0.3);
}
.sblocco-filters {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: center;
  margin-bottom: 10px;
}
.sblocco-filters label {
  display: flex;
  align-items: center;
  gap: 4px;
  font-size: 14px;
}
.sblocco-filters input[type="date"], .sblocco-filters select, .sblocco-filters input[type="text"] {
  padding: 4px 6px;
  border: 1px solid #ccc;
  border-radius: 4px;
  font-size: 13px;
}
.sblocco-filters button {
  padding: 6px 10px;
  border: none;
  border-radius: 4px;
  background: #607D8B;
  color: #fff;
  cursor: pointer;
  font-size: 13px;
}
.sblocco-summary {
  margin-bottom: 8px;
  font-size: 14px;
}
.sblocco-table-container {
  /* Rendere la tabella degli sblocchi CQ/QA ben visibile e ampia */
  max-height: 70vh;
  min-height: 50vh;
  overflow-x: auto;
  overflow-y: auto;
  border: 1px solid #ddd;
  padding-bottom: 60px; /* ampio margine inferiore per non sovrapporre la barra di scorrimento ai dati */
}
.sblocco-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}
.sblocco-table th, .sblocco-table td {
  border: 1px solid #ccc;
  padding: 4px 6px;
  text-align: left;
  white-space: nowrap;
}
/* Larghezze minime per le colonne della tabella Sblocchi CQ/QA
   1: Data/Ora    – fino a 15 caratteri (numeri e separatori)
   2: OV          – circa 5/6 numeri
   3: OP          – circa 5/6 numeri
   4: Codice      – circa 5/6 caratteri
   5: Descrizione – testo più lungo
   6: Lotto       – circa 8/9 numeri
   7: Quantità    – numero breve
   8: Stato       – fino a 10 lettere
   Queste larghezze rendono la tabella più leggibile e rispecchiano le
   specifiche richieste dall’utente.  Le colonne non menzionate in
   questo elenco (es. Descrizione) mantengono il valore di default
   definito in altre regole.*/
.sblocco-table th:nth-child(1), .sblocco-table td:nth-child(1) { min-width: 160px; }
.sblocco-table th:nth-child(2), .sblocco-table td:nth-child(2) { min-width: 80px; }
.sblocco-table th:nth-child(3), .sblocco-table td:nth-child(3) { min-width: 80px; }
.sblocco-table th:nth-child(4), .sblocco-table td:nth-child(4) { min-width: 90px; }
.sblocco-table th:nth-child(5), .sblocco-table td:nth-child(5) { min-width: 200px; }
.sblocco-table th:nth-child(6), .sblocco-table td:nth-child(6) { min-width: 100px; }
.sblocco-table th:nth-child(7), .sblocco-table td:nth-child(7) { min-width: 80px; }
.sblocco-table th:nth-child(8), .sblocco-table td:nth-child(8) { min-width: 120px; }
.sblocco-table th {
  background: #f5f5f5;
}

</style>


</head>
<body>
<div id="loginOverlay" class="login-overlay">
        <div class="login-container">
            <h2>Accesso Riservato</h2>
            <p>Inserisci il codice di accesso per continuare.</p>
            <input type="password" id="passwordInput" placeholder="Codice...">
            <button id="loginBtn">Accedi</button>
            <p id="loginError">Codice non valido. Riprova.</p>
        </div>
    </div>
    <div class="container">
        <header>
            <table class="header-layout-table">
                <tr>
                    <td rowspan="2" class="header-logo-cell">
                        <img src="https://placehold.co/100x100/f4f7f6/333?text=LOGO" alt="Logo IRA">
                    </td>
                    <td class="header-top-center-cell">
                        <p class="header-text-large">I.R.A.</p>
                        <p class="header-text-small">ISTITUTO RICERCHE APPLICATE S.p.A.</p>
                    </td>
                    <td class="header-top-right-cell">
                        <p class="header-text-large">PROGRAMMA DELLA PRODUZIONE</p>
                    </td>
                </tr>
                <tr>
                    <td class="header-bottom-center-cell">
                        <p class="header-text-large">MODULO</p>
                    </td>
                    <td class="header-bottom-right-cell">
                        <p class="header-text-small">MD 7.5-K res. 03/01/2019</p>
                        <p class="header-text-small">Pagina 1 di 1</p>
                    </td>
                </tr>
            </table>
            <div class="header-bottom-info">
                <div class="company-placeholder">
                    <p>Nome Azienda: <span id="companyName"></span></p>
                    <p>Modulo di Programmazione: <span id="programModule"></span></p>
                </div>
                <div class="date-info">
                    <p>Data Inserimento: <span id="currentDate"></span></p>
                    <p>Settimana dell'Anno: <span id="currentWeek"></span></p>
                </div>
            </div>
        </header>

        <!-- Inizio navigazione rapida: consente di raggiungere velocemente le sezioni principali
             della pagina (Gantt Produzione, Gantt Spedizioni, Programma Arrivi, Merce non Arrivata,
             Merce in Quarantena).  I link non cambiano l'URL, ma scrollano dolcemente alla
             sezione corrispondente tramite scrollIntoView.  L'identificativo del target è
             specificato nel data-scroll-target su ciascun link. -->
        <!-- Navigazione rapida: link cliccabili che permettono di raggiungere
             velocemente le sezioni principali senza dover scorrere manualmente.
             Ogni link contiene un attributo data-scroll-target che indica
             l'id del contenitore di destinazione.  Il comportamento di
             scrolling viene implementato in un listener a fine pagina. -->
        <nav id="quickNav" class="quick-navigation" style="margin-bottom: 15px; font-size: 0.95em;">
            <a href="#" data-scroll-target="ganttChartContainer">Gantt Produzione</a>
            <span style="margin: 0 4px;">|</span>
            <a href="#" data-scroll-target="warehouseGanttChartContainer">Gantt Spedizioni</a>
            <span style="margin: 0 4px;">|</span>
            <a href="#" data-scroll-target="arrivalScheduleContainer">Programma Arrivi</a>
            <span style="margin: 0 4px;">|</span>
            <a href="#" data-scroll-target="overdueArrivalsContainer">Merce non Arrivata</a>
            <span style="margin: 0 4px;">|</span>
            <a href="#" data-scroll-target="quarantineContainer">Merce in Quarantena</a>
        </nav>

        <!-- Barra di navigazione verticale numerata: consente di saltare rapidamente alle principali sezioni della pagina. La colonna rimane fissata sul lato destro dello schermo. -->
        <div id="quickNavVertical" class="quick-nav-vertical">
            <!-- Pulsante di fine sessione (logout) sempre visibile.  Il simbolo di accensione 
                 comunica chiaramente all'utente la funzione di uscita.  La struttura segue 
                 quella degli altri collegamenti (numero + tooltip) per coerenza visiva. -->
            <a id="logoutNavBtn" href="#" class="logout-nav-btn" title="Termina sessione">
                <span>&#x23FB;</span>
                <span class="tooltip">Logout</span>
            </a>
            <a href="#" data-scroll-target="ganttChartContainer"><span>1</span><span class="tooltip">Gantt Produzione</span></a>
            <a href="#" data-scroll-target="warehouseGanttChartContainer"><span>2</span><span class="tooltip">Gantt Spedizioni</span></a>
            <a href="#" data-scroll-target="arrivalScheduleContainer"><span>3</span><span class="tooltip">Programma Arrivi</span></a>
            <a href="#" data-scroll-target="overdueArrivalsContainer"><span>4</span><span class="tooltip">Merce non Arrivata</span></a>
            <a href="#" data-scroll-target="quarantineContainer"><span>5</span><span class="tooltip">Merce in Quarantena</span></a>
            <a href="#" data-scroll-target="dailyProductionContainer"><span>6</span><span class="tooltip">Produzione Giornaliera</span></a>
            <a href="#" data-scroll-target="medicalDeviceProductionContainer"><span>7</span><span class="tooltip">Produzione Dispositivi Medici</span></a>
            <a href="#" data-scroll-target="shippingScheduleContainer"><span>8</span><span class="tooltip">Programma Spedizioni</span></a>
            <a href="#" data-scroll-target="analisiContainer"><span>9</span><span class="tooltip">Analisi CQ/QA</span></a>
        </div>

        <div id="stickyControlsWrapper" class="sticky-controls-wrapper">
            <nav class="actions">
                <button id="addRowBtn" class="action-button add">Aggiungi Riga</button>
                <button id="duplicateRowBtn" class="action-button duplicate">Duplica Riga Selezionata</button>
                <button id="deleteRowBtn" class="action-button delete">Elimina Riga Selezionata</button>
                <button id="saveDataBtn" class="action-button save">Salva Dati</button>
                <button id="loadDataBtn" class="action-button load">Carica Dati</button>
                <button id="manualRefreshBtn" class="action-button load" style="background-color: #03A9F4;">Ricarica Dati da Server</button>
                <button id="importPPBtn" class="action-button import">Importa PP</button><span id="lastImportPP" class="last-import-time"></span>
                <button id="exportDataBtn" class="action-button export">Esporta Dati (CSV)</button>
                <button id="sendEmailBtn" class="action-button email">Invia via Mail</button>
                <input type="file" id="fileInput" accept=".xls, .xlsx, .csv" style="display: none;">
                <button id="logbookBtn" class="action-button primary">
                    <svg style="width:20px;height:20px;margin-right: 5px;vertical-align: middle;" viewBox="0 0 24 24">
                        <path fill="currentColor" d="M19 2H6C4.89 2 4 2.89 4 4V20C4 21.11 4.89 22 6 22H19C20.11 22 21 21.11 21 20V4C21 2.89 20.11 2 19 2M15 15H11C10.45 15 10 14.55 10 14S10.45 13 11 13H15C15.55 13 16 13.45 16 14S15.55 15 15 15M17 9H7C6.45 9 6 8.55 6 8S6.45 7 7 7H17C17.55 7 18 7.45 18 8S17.55 9 17 9M17 12H7C6.45 12 6 11.55 6 11S6.45 10 7 10H17C17.55 10 18 10.45 18 11S17.55 12 17 12M17 18H7C6.45 18 6 17.55 6 17S6.45 16 7 16H17C17.55 16 18 16.45 18 17S17.55 18 17 18Z" />
                    </svg>
                    Logbook
                </button>
                <button id="printLogbookBtn" class="action-button" style="background-color: #90A4AE; color: white;">Stampa Logbook</button>
            </nav>

            <div class="search-filter-controls">
                <input type="text" id="searchInput" placeholder="Cerca parola o numero...">
                <button id="findBtn">Trova</button>
                <button id="findNextBtn" style="display:none;">Successivo</button>
                
                <select id="filterColumn1" style="max-width: 150px;">
                    <option value="">Filtra colonna 1...</option>
                    <option value="codice">Codice</option>
                    <option value="cliente">Cliente</option>
                    <option value="prodotto">Prodotto</option>
                    <option value="operatore">Operatore</option>
                    <option value="produzioneData">Data di Produzione</option>
                    <option value="dataConfezionamento">Data di Confezionamento</option>
                    <option value="medicalDevices">Solo Dispositivi Medici</option>
                </select>
                <input type="text" id="filterValue1" placeholder="Valore filtro 1...">

                <select id="filterColumn2" style="max-width: 150px;">
                    <option value="">Filtra colonna 2...</option>
                    <option value="codice">Codice</option>
                    <option value="cliente">Cliente</option>
                    <option value="prodotto">Prodotto</option>
                    <option value="operatore">Operatore</option>
                    <option value="produzioneData">Data di Produzione</option>
                    <option value="dataConfezionamento">Data di Confezionamento</option>
                    <option value="medicalDevices">Solo Dispositivi Medici</option>
                </select>
                <input type="text" id="filterValue2" placeholder="Valore filtro 2...">

                
                <button id="clearFilterBtn">Cancella Filtri</button>
            </div>
            <div style="text-align: center; margin: 10px 0; padding: 10px; background-color: #fff3cd; border: 1px solid #ffeeba; border-radius: 8px; color: #856404; font-weight: 500;">
                <p style="margin:0; font-size: 0.9em;">
                    <strong style="color: #664d03;">Suggerimento per la stampa:</strong> Per un'ottima impaginazione, prima di stampare, premi <kbd>Ctrl</kbd> + <kbd>-</kbd> (o <kbd>Cmd</kbd> + <kbd>-</kbd> su Mac) per impostare lo zoom del browser al 67%.
                </p>
            </div>
            <div style="text-align: left; margin-bottom: 5px;">
                <p style="font-size: 1.1em; font-weight: bold; color: #333; margin: 0;">
                    Ultimo Aggiornamento File: <span id="lastModifiedTimestamp">N/A</span>
                </p>
            </div>
            <!-- Area riassuntiva per le date/ore degli ultimi import.  -->
            <!-- Questa sezione viene popolata dinamicamente e non sposta i pulsanti. -->
            <div id="lastImportsSummary" class="last-import-summary" style="font-size:0.8em;color:#333;margin-bottom:15px;"></div>
            <div id="performanceGaugeWrapper" style="display:none; margin: 10px 0; padding: 10px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; font-size:0.85em;">
                <div style="font-weight: 600; margin-bottom: 4px;">Carico del sistema</div>
                <!-- Tachimetro a semicirconferenza con puntatore rotante. Le sezioni colorate indicano le soglie di utilizzo: verde (basso carico), giallo (medio carico) e rosso (alto carico). Il puntatore ruota da -90° (0%) a +90° (100%). -->
                <div id="gaugeContainer" style="display:flex; align-items:center;">
                    <div id="performanceGauge" style="position:relative; width:160px; height:80px; border-radius:160px 160px 0 0; overflow:hidden; background: conic-gradient(#28a745 0% 30%, #ffc107 30% 60%, #dc3545 60% 100%);">
                        <div id="gaugePointer" style="position:absolute; left:50%; bottom:0; width:3px; height:80%; background:#333; transform-origin: bottom; transform: rotate(-90deg); transition: transform 0.5s ease;"></div>
                    </div>
                    <div id="performanceGaugeLabel" style="margin-left:10px;">Calcolo in corso...</div>
                </div>
                <div id="perfMetrics" style="margin-top:6px; font-size:0.8em;"></div>
            </div>
            <div class="table-header-controls">
                <h2>Dettaglio Produzione</h2>
                <div class="scroll-buttons-wrapper">
                    <button id="scrollLeftBtn" class="scroll-button">&lt;</button>
                    <button id="scrollRightBtn" class="scroll-button">&gt;</button>
                </div>
            </div>
        </div>
        <div class="table-container">
            <table id="productionTable">
                <thead>
                    <tr>
                        <th rowspan="2" class="sticky-col-header-checkbox"></th>
        <th rowspan="2" class="col-codice">Codice</th>
        <th rowspan="2" class="col-prodotto">Prodotto</th>
        <th rowspan="2" class="col-cliente">Cliente</th>
        <th rowspan="2" class="col-qty-richiesta">Quantità<br>Richiesta (Kg)</th>
        <th rowspan="2" class="col-giacenza">Giacenza<br>Magazzino (Kg)</th>
        <th rowspan="2" class="col-qty-da-produrre">Quantità<br>da Produrre (Kg)</th>
        <th rowspan="2" class="col-materie-prime">Materie<br>Prime</th>
        <th rowspan="2" class="col-macchinari">Macchinari</th>
        <th rowspan="2" class="col-operatore">Operatore</th>
        <th colspan="2" class="col-confez-group">Confezionamento<br>Richiesto</th>
        <th rowspan="2" class="col-prod-data">Data di<br>Produzione</th>
        <th rowspan="2" class="col-giorni-produzione">Giorni di<br>Produzione</th>
        <th rowspan="2" class="col-data-confez">Data di<br>Confezionamento</th>
        <th rowspan="2" class="col-cod-confez">Codice<br>Confezionamento</th>
        <th rowspan="2" class="col-lotto-sc">Lotto<br>SC</th>
        <th rowspan="2" class="col-materiale-confez">Materiale<br>Confezionamento</th>
        <th rowspan="2" class="col-data-sped">Data di<br>Spedizione</th>
        <th rowspan="2" class="col-note">Note</th>
                    </tr>
                    <tr>
                        <th class="col-confez-pezzi">Numero<br>Pezzi</th>
                        <th class="col-confez-kg-pezzo">Kg/<br>Pezzo</th>
                    </tr>
                </thead>
                <tbody>
                </tbody>
            </table>
        </div>

<div id="opiContainer" class="daily-production-container">
    <h2>Tabella OPI (Ordini di Produzione Interni)</h2>
    <div class="daily-production-controls">
    <label for="opiStartDate">Data Produzione da:</label>
    <input type="text" id="opiStartDate" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
    <label for="opiEndDate">a:</label>
    <input type="text" id="opiEndDate" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
    <label for="opiScadStartDate">Scadenza Lotto da:</label>
    <input type="text" id="opiScadStartDate" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
    <label for="opiScadEndDate">a:</label>
    <input type="text" id="opiScadEndDate" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
    <span class="control-group-separator"></span>
    <input type="text" id="filterOpiOP" placeholder="Filtra OP...">
    <input type="text" id="filterOpiOV" placeholder="Filtra OV...">
    <input type="text" id="filterOpiCodice" placeholder="Filtra Codice...">
    <input type="text" id="filterOpiArticolo" placeholder="Filtra Articolo...">
    <input type="text" id="filterOpiCliente" placeholder="Filtra Cliente...">
    <input type="text" id="filterOpiLotto" placeholder="Filtra Lotto...">
    <input type="text" id="filterOpiQuantita" placeholder="Filtra Quantità...">
    <input type="text" id="filterOpiUM" placeholder="Filtra UM...">
    <input type="text" id="filterOpiOperatore" placeholder="Filtra Operatore...">
    <button id="clearOpiFiltersBtn" class="action-button delete">Reset Filtri</button>
    <button id="importOpiBtn" class="action-button import">Importa OPI</button><span id="lastImportOPI" class="last-import-time"></span>
    <!-- Nuovo bottone per importare il file DeviceRef (informazioni specifiche per dispositivi/medicali) -->
    <button id="importDeviceRefBtn" class="action-button import">Importa DeviceRef</button><span id="lastImportDeviceRef" class="last-import-time"></span>
    <button id="sendOpiEmailBtn" class="action-button email" disabled>Invia Mail</button>
</div>
    <div class="daily-production-table-wrapper" style="max-height: 400px;">
        <table id="opiTable">
            <thead>
    <tr>
        <th>Data Produzione</th>
        <th>OP</th>
        <th>OV</th>
        <th>Codice</th>
        <th>Articolo</th>
        <th>Cliente</th>
        <th>Lotto</th>
        <th>Quantità</th>
        <th>UM</th>
        <th>Operatore</th>
        <th>Scadenza Lotto</th>
        <th>🔎</th>
    </tr>
</thead>
            <tbody>
            </tbody>
        </table>
    </div>
</div>

        <div id="salesOrderContainer" class="sales-order-container">
            <h2>Ordine di Vendita</h2>
            <div class="sales-order-controls">
                <button id="addSalesOrderRowBtn" class="action-button add">Aggiungi Riga</button>
                <button id="duplicateSalesOrderRowBtn" class="action-button duplicate">Duplica Riga Selezionata</button>
                <button id="deleteSalesOrderRowBtn" class="action-button delete">Elimina Riga Selezionata</button>
                <button id="importOVBtn" class="action-button import">Importa OV</button><span id="lastImportOV" class="last-import-time"></span>
                <button id="sendEmailOVBtn" class="action-button email">Invia via Mail</button>
            </div>
            <div class="sales-order-table-wrapper">
                <table id="salesOrderTable" class="sales-order-table">
                    <thead>
                        <tr>
                            <th></th>
                            <th class="col-ov-flag"></th>
                            <th class="col-ov">OV</th>
                            <th class="col-ov-codice">Codice</th>
                            <th class="col-ov-descrizione">Descrizione</th>
                            <th class="col-ov-quantita">Quantità Ordine</th>
                            <th class="col-ov-um">UM</th>
                            <th class="col-ov-data-consegna">Data Consegna</th>
                            <th class="col-ov-data-richiesta-cliente">Data Richiesta Cliente</th>
                            <th class="col-ov-data-conferma">Data Conferma</th>
                            <th class="col-ov-note">Note</th>
                        </tr>
                    </thead>
                    <tbody>
                    </tbody>
                </table>
            </div>
        </div>


<div id="medicalDeviceProductionContainer" class="daily-production-container">
    <h2>Produzione Medical Device</h2>
    <div class="daily-production-controls">
        <label for="medicalDeviceStartDate">Data Programma da:</label>
        <input type="text" id="medicalDeviceStartDate" class="datepicker" placeholder="gg/mm/aaaa">
        <label for="medicalDeviceEndDate">a:</label>
        <input type="text" id="medicalDeviceEndDate" class="datepicker" placeholder="gg/mm/aaaa">
        <button id="clearMedicalDeviceDateBtn" class="action-button delete">Cancella Date</button>

        <span class="control-group-separator"></span>

        <input type="text" id="filterMedicalDeviceCodice" placeholder="Filtra Codice...">
        <input type="text" id="filterMedicalDeviceDescrizione" placeholder="Filtra Descrizione...">
        <input type="text" id="filterMedicalDeviceCliente" placeholder="Filtra Cliente...">
        <!-- Nuovi filtri live per data e lotto -->
        <input type="text" id="filterMedicalDeviceData" placeholder="Filtra Data...">
        <input type="text" id="filterMedicalDeviceLotto" placeholder="Filtra Lotto...">

        <button id="clearMedicalDeviceFiltersBtn">Cancella Filtri Testo</button>
        
        <span class="control-group-separator"></span>
        <button id="addMedicalDeviceRowBtn" class="action-button add">Aggiungi Riga Manuale</button>

        <!-- Bottone per importare dati di produzione dei dispositivi medici.  Quando viene
             importato un file, la relativa data e ora vengono registrate e mostrano
             nella sezione "Ultimo Aggiornamento".  Il file importato viene inoltre
             salvato sul server e reso disponibile agli altri utenti. -->
        <span class="control-group-separator"></span>
        <button id="importMedicalProductionBtn" class="action-button import">Importa Produzione MD</button>
        <span id="lastImportMedicalProduction" class="last-import-time"></span>
    </div>
    
    <div class="daily-production-table-wrapper" style="max-height: 500px;">
        <table id="medicalDeviceProductionTable" class="daily-production-table">
            <thead>
                <tr>
                    <th style="width: 10%;">Data</th>
                    <th style="width: 15%;">Codice</th>
                    <th style="width: 25%;">Descrizione</th>
                    <th style="width: 15%;">Cliente</th>
                    <th style="width: 10%;">Lotto</th>
                    <!-- Colonna quantità: mostra i pezzi direttamente senza suffisso -->
                    <th style="width: 20%;">Quantità (pz)</th>
                    <!-- Nuova colonna per il numero teorico di scatoloni (arrotondato per eccesso).  -->
                    <th style="width: 15%;">Scatoloni (ip.)</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        </table>
    </div>
</div>

        <div id="ganttChartContainer" class="gantt-chart-container">
            <!-- Aggiornato da 14 a 30 giorni per allineare la visualizzazione al nuovo intervallo di date -->
            <h2>Grafico di Gantt (Prossimi 30 Giorni)</h2>
            <div id="ganttChart" class="gantt-chart">
            </div>
            <div id="genericTooltip" class="generic-tooltip"></div>
            <!-- Avviso ADR: viene reso visibile da JavaScript quando vengono
                 rilevate spedizioni ADR nelle due settimane programmate.  Il
                 contenuto (paragrafo e lista) viene popolato dinamicamente in
                 funzione delle spedizioni ADR rilevate.  Due pulsanti
                 (Posponi e OK) consentono all'utente di rimandare o
                 confermare l'avviso. -->
            <div id="adrNotification">
                <!-- Pulsante di chiusura: consente all'utente di nascondere
                     immediatamente il pop‑up ADR indipendentemente dal
                     ruolo o dallo stato dei dati.  È posizionato in alto a
                     destra del pannello. -->
                <span class="adr-close-btn" title="Chiudi">×</span>
                <!-- Contenitore scrollabile: il testo e la lista delle
                     spedizioni ADR vengono racchiusi in questo wrapper
                     per consentire lo scroll verticale senza perdere
                     l'accesso ai pulsanti di azione. -->
                <div class="adr-alert-content">
                    <p></p>
                    <!-- La lista delle spedizioni ADR verrà inserita qui da
                         checkAndNotifyADR().  L'elemento UL rimane vuoto
                         all'avvio e sarà popolato dinamicamente con gli
                         elenchi delle spedizioni ADR rilevate. -->
                    <ul></ul>
                </div>
                <!-- Contenitore dei pulsanti: sempre visibile in fondo
                     al pop‑up indipendentemente dalla lunghezza del
                     contenuto sopra. -->
                <div class="adr-alert-buttons">
                    <button id="adrPostponeBtn" type="button">Posponi</button>
                    <button id="adrAcknowledgeBtn" type="button">OK ho capito</button>
                </div>
            </div>

            <!-- Avviso CQ/QA: mostra gli aggiornamenti di stato da parte del
                 controllo qualità (CQ) e quality assurance (QA).  È simile
                 all'avviso ADR ma utilizza uno stile neutro.  Il
                 contenuto viene popolato dinamicamente da
                 checkAndNotifyQuality().  Due pulsanti (Posponi e OK)
                 permettono all'utente di rinviare o confermare
                 l'avviso. -->
            <div id="qualityNotification" style="display:none;">
                <span class="quality-close-btn" title="Chiudi">×</span>
                <div class="quality-alert-content">
                    <p></p>
                    <ul></ul>
                </div>
                <div class="quality-alert-buttons">
                    <button id="qualityPostponeBtn" type="button">Posponi</button>
                    <button id="qualityAcknowledgeBtn" type="button">OK Ho capito</button>
                </div>
            </div>

            <!-- Avviso Magazzino: notifica gli utenti CQ che una riga di arrivo è stata evasa
                 e spostata nella quarantena.  La struttura è simile agli avvisi ADR e QA.
                 Il contenuto viene popolato dinamicamente dalla funzione checkAndNotifyWarehouse(). -->
            <div id="warehouseNotification" style="display:none;">
                <span class="warehouse-close-btn" title="Chiudi">×</span>
                <div class="warehouse-alert-content">
                    <p></p>
                    <ul></ul>
                </div>
                <div class="warehouse-alert-buttons">
                    <button id="warehousePostponeBtn" type="button">Posponi</button>
                    <button id="warehouseAcknowledgeBtn" type="button">OK Ho capito</button>
                </div>
            </div>

            <!-- Avviso Spedizioni: notifica agli utenti con permesso spedizioni che un ordine di spedizione è stato marcato come spedito o ripristinato. -->
            <div id="shippingNotification" style="display:none;">
                <span class="shipping-close-btn" title="Chiudi">×</span>
                <div class="shipping-alert-content">
                    <p></p>
                    <ul></ul>
                </div>
                <div class="shipping-alert-buttons">
                    <button id="shippingPostponeBtn" type="button">Posponi</button>
                    <button id="shippingAcknowledgeBtn" type="button">OK Ho capito</button>
                </div>
            </div>
        </div>

         <div id="dailyProductionContainer" class="daily-production-container">
            <h2>Programma giornaliero di produzione</h2>
            <div id="print-header-info" style="display: none;">
                <h3 style="text-align: center; font-size: 14pt; margin: 0 0 5px 0;">
                    Data Programma: <span id="print-program-date"></span>
                </h3>
                <!-- Mostra il filtro corrente (se applicato) durante la stampa del programma giornaliero -->
                <div id="print-filter-info" style="text-align: center; font-size: 12pt; margin-bottom: 10px;"></div>
            </div>
            
            <div class="daily-production-controls">
                <label for="dailyProductionDateInput">Data Programma:</label>
                <input type="text" id="dailyProductionDateInput" class="datepicker" placeholder="Seleziona data">
                <button id="clearDailyProductionDateBtn" class="action-button delete">Cancella Data</button>

                <span class="control-group-separator"></span>

                <select id="filterDailyColumn" style="max-width: 150px;">
                    <option value="">Filtra colonna...</option>
                    <option value="codice">Codice</option>
                    <option value="prodotto">Prodotto</option>
                    <option value="cliente">Cliente</option>
                    <option value="operatore">Operatore</option>
                </select>
                <input type="text" id="filterDailyValue" placeholder="Valore filtro...">
                <!-- Datalist per suggerimenti operatore (popolato dinamicamente) -->
                <datalist id="operatorSuggestionsList"></datalist>
                <!-- Datalist per suggerimenti macchinari nel programma giornaliero -->
                <datalist id="macchinariOptionsListDaily"></datalist>
               
                <button id="clearDailyFilterBtn">Cancella Filtro</button>
                
                <span class="control-group-separator"></span>

                <button id="addDailyRowBtn" class="action-button add">Aggiungi Riga Vuota</button>
                <button id="duplicateDailyRowBtn" class="action-button duplicate">Duplica Riga Selezionata</button>
                <button id="deleteDailyRowBtn" class="action-button delete">Elimina Riga Selezionata</button>
                
                <span class="control-group-separator"></span>

                <button id="saveDailyDataBtn" class="action-button save">Salva Programma</button>
                <button id="loadDailyDataBtn" class="action-button load">Carica Programma</button>
                <button id="exportDailyExcelBtn" class="action-button export">Esporta Excel</button>
                <button id="exportDailyPdfBtn" class="action-button export">Esporta PDF</button>
                <button id="exportDailyWordBtn" class="action-button export">Esporta Word</button>
            </div>
            
            <div class="daily-production-table-wrapper">
                <table id="dailyProductionTable" class="daily-production-table">
                    <thead>
                        <tr>
                            <th></th>
                            <!-- Prima colonna: OP (nuovo nome per OPE) -->
                            <th class="col-daily-op">OP</th>
                            <th class="col-daily-ov">OV</th>
                            <th class="col-daily-codice">Codice</th>
                            <th class="col-daily-prodotto">Prodotto</th>
                            <th class="col-daily-cliente">Cliente</th>
                            <!-- Spostiamo la colonna Lotto immediatamente dopo Cliente -->
                            <th class="col-daily-lotto">Lotto</th>
                            <th class="col-daily-quantita">Quantità</th>
                            <th class="col-daily-macchinario">Macchinario</th>
                            <!-- Diminuiamo leggermente la larghezza della colonna di quantità di confezionamento -->
                            <th class="col-daily-quantita-confez">Quantità<br>Confezionamento</th>
                            <th class="col-daily-operazioni">Operazioni</th>
                            <th class="col-daily-operatori">Operatori</th>
                            <th class="col-daily-esito">Esito</th>
                            <th class="col-daily-qty-prodotta">Quantità<br>Prodotta</th>
                            <th class="col-daily-data-avallo">Data Avallo</th>
                        </tr>
                    </thead>
                    <tbody>
                    </tbody>
                </table>
            </div>
        </div>
        
<div id="shippingScheduleContainer" class="daily-production-container">
    <h2>Programma Giornaliero di Spedizione</h2>
    <div class="daily-production-controls">
        <label for="shippingStartDate">Data Spedizione da:</label>
        <input type="text" id="shippingStartDate" class="datepicker" placeholder="gg/mm/aaaa">
        <label for="shippingEndDate">a:</label>
        <input type="text" id="shippingEndDate" class="datepicker" placeholder="gg/mm/aaaa">
        <button id="clearShippingDateBtn" class="action-button delete">Cancella Date</button>
        
        <span class="control-group-separator"></span>

        <select id="filterShippingColumn" style="max-width: 150px;">
    <option value="">Filtra colonna...</option>
    <option value="ov">OV</option>
    <option value="codiceArticolo">Codice Articolo</option>
    <option value="descrizioneArticolo">Descrizione Articolo</option>
    <option value="ragioneSociale">Cliente</option>
</select>

        <input type="text" id="filterShippingValue" placeholder="Valore filtro...">
       
        <button id="clearShippingFilterBtn">Cancella Filtro</button>

        <span class="control-group-separator"></span>

        <button id="addShippingRowBtn" class="action-button add">Aggiungi Riga</button>
        <button id="duplicateShippingRowBtn" class="action-button duplicate">Duplica Riga</button>
        <button id="deleteShippingRowBtn" class="action-button delete">Elimina Riga</button>
        
        <span class="control-group-separator"></span>

        <button id="importOSBtn" class="action-button import">Importa OS</button><span id="lastImportOS" class="last-import-time"></span>
        <button id="exportShippingDataBtn" class="action-button export">Esporta CSV</button>
        <button id="printShippingBtn" class="action-button" style="background-color: #B0BEC5; color: white;">Stampa OS</button> 
        <button id="sendShippingEmailBtn" class="action-button email">Invia Mail</button>
    </div>
    <div class="daily-production-table-wrapper" style="max-height: 400px;">
        <table id="shippingScheduleTable">
            <thead>
                <tr>
                    <th></th>
                    <th>OV</th>
                    <th>Codice Articolo</th>
                    <th>Descrizione Articolo</th>
                    <th>Quantità</th>
                    <th>UM</th>
                    <th>Data Consegna</th>
                    <th>Data Conferma</th>
                    <th>Ragione Sociale</th>
                    <th>Rif. Cliente</th>
                    <th>Indirizzo</th>
                    <th>CAP</th>
                    <th>Città</th>
                    <th>Provincia</th>
                    <th>Telefono</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        </table>
    </div>
</div>


<!-- Barra di scorrimento esterna e pulsanti laterali per il Gantt Spedizioni -->
<div id="warehouseGanttExternalScrollbar" class="gantt-external-scrollbar">
    <div class="gantt-external-sizer" style="height:1px;"></div>
</div>
<div id="warehouseGanttScrollButtonsWrapper" class="gantt-scroll-buttons-wrapper">
    <button id="warehouseGanttScrollLeftBtn" class="scroll-button">&lt;</button>
    <button id="warehouseGanttScrollRightBtn" class="scroll-button">&gt;</button>
</div>

<div id="arrivalScheduleContainer" class="daily-production-container">
    <h2>Programma Giornaliero di Arrivo Merce</h2>
    <div class="daily-production-controls">
        <label for="arrivalStartDate">Data Arrivo da:</label>
        <input type="text" id="arrivalStartDate" class="datepicker" placeholder="gg/mm/aaaa">
        <label for="arrivalEndDate">a:</label>
        <input type="text" id="arrivalEndDate" class="datepicker" placeholder="gg/mm/aaaa">
        <button id="clearArrivalDateBtn" class="action-button delete">Cancella Date</button>
        
        <span class="control-group-separator"></span>

        <select id="filterArrivalColumn" style="max-width: 150px;">
    <option value="">Filtra colonna...</option>
    <!-- Usa "OA" come etichetta visibile pur mantenendo il valore "ov" per compatibilità con il codice esistente -->
    <option value="ov">OA</option>
    <option value="codiceArticolo">Codice Articolo</option>
    <option value="descrizioneArticolo">Descrizione Articolo</option>
    <option value="ragioneSociale">Cliente</option>
</select>

        <input type="text" id="filterArrivalValue" placeholder="Valore filtro...">
       
        <button id="clearArrivalFilterBtn">Cancella Filtro</button>

        <span class="control-group-separator"></span>

        <button id="addArrivalRowBtn" class="action-button add">Aggiungi Riga</button>
        <button id="duplicateArrivalRowBtn" class="action-button duplicate">Duplica Riga</button>
        <button id="deleteArrivalRowBtn" class="action-button delete">Elimina Riga</button>
        
        <span class="control-group-separator"></span>

        <button id="importArrivalsBtn" class="action-button import">Importa Arrivi</button><span id="lastImportArrivals" class="last-import-time"></span>                 
        <div class="file-status-group" style="position:relative;">
        <button id="importLayoutBtn" class="action-button" style="background-color: #B2EBF2; color: #00796B;">Layout</button><span id="lastImportLayout" class="last-import-time"></span>
        <span id="layoutFileStatus" class="file-status-flag" style="display:none;">✔</span>
        </div>
        <button id="exportPropostaLayoutBtn" class="action-button" style="background-color: #009688; color: white;">Proposta Layout (PDF)</button>

        
        <button id="exportArrivalDataBtn" class="action-button export">Esporta CSV</button>
        <button id="sendArrivalEmailBtn" class="action-button email">Invia Mail</button>
    </div>
    <div class="daily-production-table-wrapper" style="max-height: 400px;">
        <table id="arrivalScheduleTable">
            <thead>
                <tr>
                    <th></th>
                    <th>OA</th>
                    <th>Codice Articolo</th>
                    <th>Descrizione Articolo</th>
                    <th>Layout</th>
                    <th>Quantità</th>
                    <th>UM</th>
                    <th>Data Consegna</th>
                    <th>Data Conferma</th>
                    <th>Ragione Sociale</th>
                    <th>Rif. Cliente</th>
                    <th>Indirizzo</th>
                    <th>CAP</th>
                    <th>Città</th>
                    <th>Provincia</th>
                    <th>Telefono</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        </table>
    </div>
</div>

        <!-- Le barre di scorrimento esterne e i pulsanti laterali verranno inseriti più sopra -->

        <div id="warehouseGanttChartContainer" class="gantt-chart-container">
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
        <h2>Grafico di Gantt Magazzino (Prossimi 30 Giorni - Spedizioni)</h2>
        <div style="display:flex; gap:8px;">
            <button id="printWarehouseGanttBtn" class="action-button" style="background-color: #78909C; color: white;">Stampa Gantt Spedizioni</button>
            <button id="packingListBtn" class="action-button" style="background-color: #FF9800; color: white;">Packing&nbsp;List</button>
            <button id="sbloccoBtn" class="action-button" style="background-color: #8E24AA; color: white;">Sblocchi CQ/QA</button>
        </div>
    </div>
    <!-- Pulsanti di scorrimento rimossi: per lo scorrimento si utilizza la barra
         orizzontale nativa del contenitore.  Vedi la sezione script per la
         rimozione della loro logica. -->
    <!-- Contenitore reale del grafico di Gantt di magazzino.  Rimuoviamo una seconda
         istanza duplicata di questo div che creava problemi di rendering.  Ora
         esiste una sola definizione di #warehouseGanttChart, dentro la quale il
         grafico viene generato dinamicamente. -->
    <div id="warehouseGanttScrollWrapper" class="gantt-inline-wrapper"><div id="warehouseGanttChart" class="gantt-chart warehouse-gantt-chart"></div></div>

    <!-- Fine del contenitore del Gantt Magazzino; la tabella della quarantena è stata spostata fuori da questo wrapper -->
    

<!-- Sezione per la merce in quarantena spostata qui per non interferire con la larghezza del Gantt -->
<div id="quarantineContainer" class="daily-production-container">
    <h2>Merce in Quarantena</h2>
    <!-- Filtri e azioni per la tabella quarantena: consentono di cercare tra le righe e cancellare selezioni -->
    <div class="daily-production-controls" id="quarantineFilters" style="margin-bottom:10px;">
        <input type="text" id="filterQuarantineOV" placeholder="Filtra OA...">
        <input type="text" id="filterQuarantineCodice" placeholder="Filtra Codice Articolo...">
        <input type="text" id="filterQuarantineDescrizione" placeholder="Filtra Descrizione Articolo...">
        <label for="filterQuarantineDataDa">Data Consegna da:</label>
        <input type="text" id="filterQuarantineDataDa" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
        <label for="filterQuarantineDataA">a:</label>
        <input type="text" id="filterQuarantineDataA" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
        <input type="text" id="filterQuarantineRagSoc" placeholder="Filtra Ragione Sociale...">
        <button id="clearQuarantineFiltersBtn" class="action-button delete">Reset Filtri</button>
        <button id="deleteQuarantineRowBtn" class="action-button delete">Elimina Riga</button>
    </div>
    <div class="daily-production-table-wrapper" style="max-height: 300px;">
        <table id="quarantineTable">
            <thead>
                <tr>
                    <th></th>
                    <th>OA</th>
                    <th>Codice Articolo</th>
                    <th>Descrizione Articolo</th>
                    <th>Layout</th>
                    <th>Quantità</th>
                    <th>UM</th>
                    <th>Data Consegna</th>
                    <th>Data Conferma</th>
                    <th>Ragione Sociale</th>
                    <th>Rif. Cliente</th>
                    <th>Indirizzo</th>
                    <th>CAP</th>
                    <th>Città</th>
                    <th>Provincia</th>
                    <th>Telefono</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        </table>
    </div>
</div>

<div id="overdueArrivalsContainer" class="daily-production-container">
    <h2>Merce non Arrivata </h2>
<div class="daily-production-controls" id="overdueArrivalsFilters" style="margin-bottom:10px;">
    <input type="text" id="filterOverdueOV" placeholder="Filtra OA...">
    <input type="text" id="filterOverdueCodice" placeholder="Filtra Codice Articolo...">
    <input type="text" id="filterOverdueDescrizione" placeholder="Filtra Descrizione Articolo...">
    <label for="filterOverdueDataDa">Data Consegna da:</label>
    <input type="text" id="filterOverdueDataDa" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
    <label for="filterOverdueDataA">a:</label>
    <input type="text" id="filterOverdueDataA" class="datepicker" placeholder="gg/mm/aaaa" style="width:120px;">
    <input type="text" id="filterOverdueRagSoc" placeholder="Filtra Ragione Sociale...">
    <button id="clearOverdueFiltersBtn" class="action-button delete">Reset Filtri</button>
</div>

    <div class="daily-production-table-wrapper" style="max-height: 300px;">
        <table id="overdueArrivalsTable">
            <thead>
                <tr>
                    <th></th>
                    <th>OA</th>
                    <th>Codice Articolo</th>
                    <th>Descrizione Articolo</th>
                    <th>Layout</th>
                    <th>Quantità</th>
                    <th>UM</th>
                    <th>Data Consegna</th>
                    <th>Data Conferma</th>
                    <th>Ragione Sociale</th>
                    <th>Rif. Cliente</th>
                    <th>Indirizzo</th>
                    <th>CAP</th>
                    <th>Città</th>
                    <th>Provincia</th>
                    <th>Telefono</th>
                </tr>
            </thead>
            <tbody>
                </tbody>
        </table>
    </div>
</div>

 
<!-- Sezione esterna: Merce in Scadenza (spostata fuori dal container Merce non Arrivata) -->
<div id="expiringGoodsContainer" class="daily-production-container">
    <h2>Merce in Scadenza</h2>
    <div class="daily-production-controls">
        <label for="expiringStartDate">Data Scadenza da:</label>
        <input type="text" id="expiringStartDate" class="datepicker" placeholder="gg/mm/aaaa">
        <label for="expiringEndDate">a:</label>
        <input type="text" id="expiringEndDate" class="datepicker" placeholder="gg/mm/aaaa">
        <button id="clearExpiringDateBtn" class="action-button delete">Cancella Date</button>

        <span class="control-group-separator"></span>

        <select id="filterExpiringColumn" style="max-width: 150px;">
            <option value="">Filtra colonna...</option>
            <option value="codice">Codice</option>
            <option value="articolo">Articolo</option>
            <option value="famiglia">Famiglia</option>
            <option value="linea">Linea</option>
        </select>
        <input type="text" id="filterExpiringValue" placeholder="Valore filtro...">
        <button id="clearExpiringFilterBtn">Cancella Filtro</button>

        <span class="control-group-separator"></span>

        <button id="addExpiringRowBtn" class="action-button add">Aggiungi Riga</button>
        <button id="duplicateExpiringRowBtn" class="action-button duplicate">Duplica Riga</button>
        <button id="deleteExpiringRowBtn" class="action-button delete">Elimina Riga</button>

        <span class="control-group-separator"></span>

        <input type="file" id="inventoryInput" accept=".xls, .xlsx, .csv" style="display:none;">
        <button id="importInventoryBtn" class="action-button import">Inventario</button><span id="lastImportInventory" class="last-import-time"></span>
        <button id="exportExpiringDataBtn" class="action-button export">Esporta CSV</button>
        <button id="sendExpiringEmailBtn" class="action-button email">Invia Mail</button>
    </div>
    <div class="daily-production-table-wrapper" style="max-height: 400px;">
        <table id="expiringGoodsTable">
            <thead>
                <tr>
                    <th></th>
                    <th>Codice</th>
                    <th>Articolo</th>
                    <th>Lotto</th>
                    <th>Scadenza</th>
                    <th>Quantità</th>
                    <th>UM</th>
                    <th>Layout</th>
                    <th>Famiglia</th>
                    <th>Linea</th>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
    </div>
</div>

         <div id="analisiContainer" class="daily-production-container">
            <h2>Analisi Chimico - Fisiche e Microbiologiche</h2>
            <div class="analisi-table-actions">


               <div class="file-status-group" style="position:relative;">
    <input type="file" id="referenzeInput" accept=".xls, .xlsx, .csv" style="display: none;">
    <button id="importReferenzeBtn" class="action-button import">Importa Referenze</button>
    <span id="referenzeFileStatus" class="file-status-flag" style="display:none;">✔</span><span id="lastImportReferenze" class="last-import-time"></span>
</div>
<div class="file-status-group" style="position:relative;">
    <input type="file" id="pianoAnaliticoInput" accept=".xls, .xlsx, .csv" style="display: none;">
    <button id="importPianoAnaliticoBtn" class="action-button import">Importa Piano Analitico</button>
    <span id="pianoAnaliticoFileStatus" class="file-status-flag" style="display:none;">✔</span><span id="lastImportPianoAnalitico" class="last-import-time"></span>
</div>
                <button id="addAnalisiRowBtn" class="action-button add">Aggiungi Riga</button>
                <button id="duplicateAnalisiRowBtn" class="action-button duplicate">Duplica Riga</button>
                <button id="deleteAnalisiRowBtn" class="action-button delete">Elimina Riga</button>
                <button id="exportAnalisiPdfBtn" class="action-button export">Esporta PDF Analisi</button>
            </div>
           <div class="daily-production-controls">
    <label for="analisiStartDate">Data Analisi da:</label>
    <input type="text" id="analisiStartDate" class="datepicker" placeholder="gg/mm/aaaa">
    <label for="analisiEndDate">a:</label>
    <input type="text" id="analisiEndDate" class="datepicker" placeholder="gg/mm/aaaa">

    <span class="control-group-separator"></span>

    <label for="searchLottoInput">Cerca per Lotto:</label>
    <input type="text" id="searchLottoInput" placeholder="Inserisci lotto...">

    <button id="clearAnalisiFilterBtn" class="action-button delete" style="margin-left: 10px;">Cancella Filtri</button>
</div>
            <div class="daily-production-table-wrapper">
                <table id="analisiTable">
                    <thead>
                        </thead>
                    <tbody>
                        </tbody>
                </table>
            </div>
        </div>

        <div id="logbookContainer" class="daily-production-container" style="display: none;">
            <h2>Logbook Attività Importazione</h2>
            <div class="daily-production-controls" style="justify-content: flex-start; gap: 5px; flex-wrap: nowrap;">
                <label>Da:</label>
                <input type="text" id="logbookStartDate" placeholder="Data" style="width: 110px; flex-shrink: 0;">
                <input type="text" id="logbookStartTime" placeholder="Ora (opz.)" style="width: 80px; flex-shrink: 0;">
                <label>A:</label>
                <input type="text" id="logbookEndDate" placeholder="Data" style="width: 110px; flex-shrink: 0;">
                <input type="text" id="logbookEndTime" placeholder="Ora (opz.)" style="width: 80px; flex-shrink: 0;">
                <button id="clearLogbookFilterBtn" class="action-button delete">Pulisci Filtri</button>
            </div>
            <div style="background-color: #f8f8f8; border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; max-height: 400px; overflow-y: auto; font-family: 'monospace'; font-size: 0.9em; line-height: 1.4;">
                <pre id="logbookContent"></pre>
            </div>
            <button id="clearLogbookBtn" class="action-button delete" style="margin-top: 15px;">Pulisci Log</button>
        </div>
    </div>


    <div id="customModal" class="modal-overlay">
        <div class="modal-content">
            <h3 id="modalTitle"></h3>
            <p id="modalMessage"></p>
            <div id="modalButtons" class="modal-buttons">
            </div>
        </div>
    </div>
    <!-- Modal per la creazione della Packing List -->
    <div id="packingListModal" class="packing-list-modal">
        <div class="packing-list-content">
            <h3>Seleziona gli Ordini di Vendita da includere</h3>
            <ul id="packingListItems" class="packing-list-items"></ul>
            <div class="packing-list-modal-footer">
                <button id="packingListCreateBtn">Crea Packing List</button>
                <button id="packingListCloseBtn">Chiudi</button>
            </div>
        </div>
    </div>

    <!-- Modal legacy per gestire e visualizzare gli sblocchi CQ/QA (id distinti per evitare conflitti) -->
<div id="legacySbloccoModal" class="sblocco-modal" style="display:none;">
        <div class="sblocco-content">
            <h3>Registro Sblocchi CQ/QA (Legacy)</h3>
            <div class="sblocco-filters">
                <label>Da: <input type="date" id="legacySbloccoStartDate"></label>
                <label>A: <input type="date" id="legacySbloccoEndDate"></label>
                <label>Stato:
                    <select id="legacySbloccoStateFilter">
                        <option value="all">Tutti</option>
                        <option value="green">Conforme (Green)</option>
                        <option value="red">Non conforme (Red)</option>
                    </select>
                </label>
                <label>Cerca: <input type="text" id="legacySbloccoSearchInput" placeholder="Cerca..."></label>
                <button id="legacySbloccoExportBtn">Esporta CSV</button>
                <button id="legacySbloccoPrintBtn">Stampa</button>
                <button id="legacySbloccoCloseBtn">Chiudi</button>
            </div>
            <div class="sblocco-summary">
                <span id="legacySbloccoTotalCount">Totale sblocchi: 0</span> |
                <span id="legacySbloccoGreenCount">Conformi: 0</span> |
                <span id="legacySbloccoRedCount">Non conformi: 0</span>
            </div>
            <div class="sblocco-table-container">
                <table id="legacySbloccoTable" class="sblocco-table">
                    <thead>
                        <tr>
                            <th>Data/Ora</th>
                            <th>OV</th>
                            <th>OP</th>
                            <th>Codice</th>
                            <th>Descrizione</th>
                            <th>Lotto</th>
                            <th>Quantità</th>
                            <th>Stato</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Nuovo modulo per gestire e visualizzare gli sblocchi CQ/QA -->
    <div id="sbloccoModal" class="sblocco-modal">
      <div class="sblocco-content">
        <h3>Registro Sblocchi CQ/QA</h3>
        <!-- Elementi legacy rimossi: i campi e la tabella legacy sono ora definiti solo nel modulo legacy. -->
        <!-- Barra comandi (sempre visibile) -->
        <div class="sblocco-filters" style="justify-content: flex-end; gap: 8px;">
          <button id="sbloccoExportBtn">Esporta CSV</button>
          <button id="sbloccoPrintBtn">Stampa</button>
          <button id="sbloccoCloseBtn">Chiudi</button>
        </div>
        <!-- SEZIONE CQ -->
        <div class="sblocco-section">
          <h4 style="margin:8px 0 6px;">Sblocchi CQ</h4>
          <div class="sblocco-filters" id="sbloccoCQFilters">
            <label>Da: <input type="date" id="sbloccoCQStartDate"></label>
            <label>A: <input type="date" id="sbloccoCQEndDate"></label>
            <label>Stato:
              <select id="sbloccoCQStateFilter">
                <option value="all">Tutti</option>
                <option value="green">Conforme (Green)</option>
                <option value="yellow">Deroga (Yellow)</option>
                <option value="red">Non conforme (Red)</option>
                <option value="white">In analisi (White)</option>
              </select>
            </label>
            <label>Codice: <input type="text" id="sbloccoCQFilterCodice" placeholder="Codice"></label>
            <label>OV: <input type="text" id="sbloccoCQFilterOV" placeholder="OV"></label>
            <label>OP: <input type="text" id="sbloccoCQFilterOP" placeholder="OP"></label>
            <label>Articolo: <input type="text" id="sbloccoCQFilterDescrizione" placeholder="Nome articolo"></label>
            <label>Lotto: <input type="text" id="sbloccoCQFilterLotto" placeholder="Lotto"></label>
            <button id="sbloccoCQResetBtn" class="action-button delete">Reset</button>
            <button id="addCQRowBtn" class="action-button add">Aggiungi riga vuota</button>
          </div>
          <div class="sblocco-summary">
            <span id="sbloccoCQTotalCount">Totale sblocchi: 0</span> |
            <span id="sbloccoCQGreenCount">Conformi: 0</span> |
            <span id="sbloccoCQRedCount">Non conformi: 0</span>
          </div>
          <div class="sblocco-table-container">
            <table id="sbloccoTableCQ" class="sblocco-table">
              <thead>
                <tr>
                  <th>Data/Ora</th>
                  <th>OV</th>
                  <th>OP</th>
                  <th>Codice</th>
                  <th>Descrizione</th>
                  <th>Lotto</th>
                  <th>Quantità</th>
                  <th>Stato</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
            <!-- Spazio per evitare sovrapposizione barra di scorrimento -->
            <div style="height:48px"></div>
          </div>
        </div>
        <!-- SEZIONE QA -->
        <div class="sblocco-section" style="margin-top:14px;">
          <h4 style="margin:8px 0 6px;">Sblocchi QA</h4>
          <div class="sblocco-filters" id="sbloccoQAFilters">
            <label>Da: <input type="date" id="sbloccoQAStartDate"></label>
            <label>A: <input type="date" id="sbloccoQAEndDate"></label>
            <label>Stato:
              <select id="sbloccoQAStateFilter">
                <option value="all">Tutti</option>
                <option value="green">Conforme (Green)</option>
                <option value="yellow">Deroga (Yellow)</option>
                <option value="red">Non conforme (Red)</option>
                <option value="white">In analisi (White)</option>
              </select>
            </label>
            <label>Codice: <input type="text" id="sbloccoQAFilterCodice" placeholder="Codice"></label>
            <label>OV: <input type="text" id="sbloccoQAFilterOV" placeholder="OV"></label>
            <label>OP: <input type="text" id="sbloccoQAFilterOP" placeholder="OP"></label>
            <label>Articolo: <input type="text" id="sbloccoQAFilterDescrizione" placeholder="Nome articolo"></label>
            <label>Lotto: <input type="text" id="sbloccoQAFilterLotto" placeholder="Lotto"></label>
            <button id="sbloccoQAResetBtn" class="action-button delete">Reset</button>
            <button id="addQARowBtn" class="action-button add">Aggiungi riga vuota</button>
          </div>
          <div class="sblocco-summary">
            <span id="sbloccoQATotalCount">Totale sblocchi: 0</span> |
            <span id="sbloccoQAGreenCount">Conformi: 0</span> |
            <span id="sbloccoQARedCount">Non conformi: 0</span>
          </div>
          <div class="sblocco-table-container">
            <table id="sbloccoTableQA" class="sblocco-table">
              <thead>
                <tr>
                  <th>Data/Ora</th>
                  <th>OV</th>
                  <th>OP</th>
                  <th>Codice</th>
                  <th>Descrizione</th>
                  <th>Lotto</th>
                  <th>Quantità</th>
                  <th>Stato</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
            <div style="height:48px"></div>
          </div>
        </div>
      </div>
    </div>
<script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
    <script src="https://cdn.jsdelivr.net/npm/flatpickr/dist/l10n/it.js"></script>
    <script>
    // === Script per la nuova gestione delle tabelle Sblocchi CQ e QA ===
    document.addEventListener('DOMContentLoaded', function() {
      const STORAGE_KEYS = { CQ: 'sbloccoData_CQ', QA: 'sbloccoData_QA' };
      function loadData(type) {
        const key = STORAGE_KEYS[type];
        try {
          const raw = localStorage.getItem(key);
          return raw ? JSON.parse(raw) : [];
        } catch (e) {
          return [];
        }
      }
      function saveData(type, data) {
        const key = STORAGE_KEYS[type];
        localStorage.setItem(key, JSON.stringify(data));
      }
      function addRow(type) {
        const data = loadData(type);
        const now = new Date();
        const timestamp = now.toISOString();
        const entry = {
          timestamp: timestamp,
          date: timestamp.split('T')[0],
          dateTime: now.toLocaleString('it-IT'),
          ov: '',
          op: '',
          codice: '',
          descrizione: '',
          lotto: '',
          quantita: '',
          state: ''
        };
        data.push(entry);
        saveData(type, data);
        renderTable(type);
      }
      function updateField(type, index, field, value) {
        const data = loadData(type);
        if (!data[index]) return;
        data[index][field] = value;
        if (!data[index].timestamp) {
          const now = new Date();
          const ts = now.toISOString();
          data[index].timestamp = ts;
          data[index].date = ts.split('T')[0];
          data[index].dateTime = now.toLocaleString('it-IT');
        }
              saveData(type, data);
              // Se si modifica il campo di stato, registra l'evento nel registro
              // degli sblocchi, includendo i valori correnti della riga.  Questo
              // permette di tracciare tutte le approvazioni CQ/QA in maniera
              // persistente.  Usa la nuova API recordSbloccoEvent se disponibile.
              if (field === 'state' && typeof window.recordSbloccoEvent === 'function') {
                const entry = data[index];
                try {
                  window.recordSbloccoEvent({
                    type: type,
                    ov: entry.ov || '',
                    op: entry.op || '',
                    codice: entry.codice || '',
                    descrizione: entry.descrizione || '',
                    lotto: entry.lotto || '',
                    quantita: entry.quantita || '',
                    state: value || ''
                  });
                } catch (err) {
                  console.warn('Errore nel tracciamento dello sblocco:', err);
                }
              }
              updateSummary(type);
      }
      function applyFilters(type, data) {
        const getVal = (id) => {
          const el = document.getElementById(id);
          return el ? el.value : '';
        };
        const startVal = getVal(`sblocco${type}StartDate`);
        const endVal = getVal(`sblocco${type}EndDate`);
        const stateVal = getVal(`sblocco${type}StateFilter`) || 'all';
        const codVal = getVal(`sblocco${type}FilterCodice`).trim().toLowerCase();
        const ovVal = getVal(`sblocco${type}FilterOV`).trim().toLowerCase();
        const opVal = getVal(`sblocco${type}FilterOP`).trim().toLowerCase();
        const descrVal = getVal(`sblocco${type}FilterDescrizione`).trim().toLowerCase();
        const lottoVal = getVal(`sblocco${type}FilterLotto`).trim().toLowerCase();
        const filtered = [];
        data.forEach((item, idx) => {
          let match = true;
          if (startVal && item.date < startVal) match = false;
          if (endVal && item.date > endVal) match = false;
          if (stateVal !== 'all' && item.state !== stateVal) match = false;
          if (codVal && !(item.codice || '').toLowerCase().includes(codVal)) match = false;
          if (ovVal && !(item.ov || '').toLowerCase().includes(ovVal)) match = false;
          if (opVal && !(item.op || '').toLowerCase().includes(opVal)) match = false;
          if (descrVal && !(item.descrizione || '').toLowerCase().includes(descrVal)) match = false;
          if (lottoVal && !(item.lotto || '').toLowerCase().includes(lottoVal)) match = false;
          if (match) filtered.push({ item: item, index: idx });
        });
        return filtered;
      }
      function updateSummary(type) {
        const data = loadData(type);
        const filtered = applyFilters(type, data);
        const total = filtered.length;
        let green = 0, red = 0;
        filtered.forEach(obj => {
          if (obj.item.state === 'green') green++;
          else if (obj.item.state === 'red') red++;
        });
        const totalSpan = document.getElementById(`sblocco${type}TotalCount`);
        const greenSpan = document.getElementById(`sblocco${type}GreenCount`);
        const redSpan = document.getElementById(`sblocco${type}RedCount`);
        if (totalSpan) totalSpan.textContent = `Totale sblocchi: ${total}`;
        if (greenSpan) greenSpan.textContent = `Conformi: ${green}`;
        if (redSpan) redSpan.textContent = `Non conformi: ${red}`;
      }
      function renderTable(type) {
        // === Inizio misurazione prestazioni ===
        const _renderStart = (typeof performance !== 'undefined' && typeof performance.now === 'function') ? performance.now() : Date.now();
        const tbody = document.querySelector(`#sbloccoTable${type} tbody`);
        if (!tbody) return;
        const data = loadData(type);
        const filtered = applyFilters(type, data);
        filtered.sort((a, b) => new Date(b.item.timestamp || 0) - new Date(a.item.timestamp || 0));
        tbody.innerHTML = '';
        filtered.forEach(({ item, index }) => {
          const tr = document.createElement('tr');
          tr.setAttribute('data-index', index);
          const tdDate = document.createElement('td');
          tdDate.textContent = item.dateTime || '';
          tr.appendChild(tdDate);
          function createInputCell(value, field) {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'text';
            input.value = value || '';
            input.style.width = '100%';
            input.dataset.field = field;
            input.dataset.type = type;
            input.dataset.index = index;
            input.addEventListener('input', function() {
              updateField(this.dataset.type, parseInt(this.dataset.index), this.dataset.field, this.value);
            });
            td.appendChild(input);
            return td;
          }
          tr.appendChild(createInputCell(item.ov, 'ov'));
          tr.appendChild(createInputCell(item.op, 'op'));
          tr.appendChild(createInputCell(item.codice, 'codice'));
          tr.appendChild(createInputCell(item.descrizione, 'descrizione'));
          tr.appendChild(createInputCell(item.lotto, 'lotto'));
          tr.appendChild(createInputCell(item.quantita, 'quantita'));
          const tdState = document.createElement('td');
          const select = document.createElement('select');
          select.dataset.field = 'state';
          select.dataset.type = type;
          select.dataset.index = index;
          const opts = [ {v:'',l:''}, {v:'green',l:'Green'}, {v:'yellow',l:'Yellow'}, {v:'red',l:'Red'}, {v:'white',l:'White'} ];
          opts.forEach(opt => {
            const o = document.createElement('option');
            o.value = opt.v;
            o.textContent = opt.l;
            select.appendChild(o);
          });
          select.value = item.state || '';
          select.addEventListener('change', function() {
            updateField(this.dataset.type, parseInt(this.dataset.index), this.dataset.field, this.value);
            renderTable(type);
          });
          tdState.appendChild(select);
          tr.appendChild(tdState);
          tbody.appendChild(tr);
        });
        const currentRows = filtered.length;
        const minRows = 5;
        if (currentRows < minRows) {
          for (let i = 0; i < (minRows - currentRows); i++) {
            const trBlank = document.createElement('tr');
            for (let j = 0; j < 8; j++) {
              const td = document.createElement('td');
              td.innerHTML = '&nbsp;';
              trBlank.appendChild(td);
            }
            tbody.appendChild(trBlank);
          }
        }
        updateSummary(type);
        // === Fine misurazione prestazioni ===
        const _renderEnd = (typeof performance !== 'undefined' && typeof performance.now === 'function') ? performance.now() : Date.now();
        try {
          if (!window.lastRenderDurations) window.lastRenderDurations = {};
          window.lastRenderDurations[type] = _renderEnd - _renderStart;
          if (typeof updatePerfMetrics === 'function') updatePerfMetrics();
          // Emette un messaggio di debug nel console per facilitare il monitoraggio
          // delle prestazioni tramite gli strumenti di sviluppo (F12).  Il
          // messaggio include il tipo di tabella e la durata dell'ultimo
          // rendering.
          try {
            const dur = window.lastRenderDurations[type];
            if (typeof console !== 'undefined' && typeof console.debug === 'function') {
              console.debug('[Performance] renderTable ' + type + ': ' + dur.toFixed(1) + ' ms');
            }
          } catch (e) {}
        } catch (e) {
          console.warn('Errore nell\'aggiornamento delle metriche di rendering:', e);
        }
      }

    /**
     * Funzione debounce: ritarda l'esecuzione della funzione fornita finché
     * non siano trascorsi `wait` millisecondi dall'ultimo invocazione.
     * Utile per ridurre il numero di rendering quando si digitano i filtri,
     * evitando di eseguire più volte il ricalcolo della tabella durante la
     * digitazione continua.
     * @param {Function} func - La funzione da eseguire con ritardo.
     * @param {number} wait - Il tempo di attesa in millisecondi.
     * @returns {Function} - Una nuova funzione che implementa il debounce.
     */
    function debounce(func, wait) {
        let timeout;
        return function () {
            const context = this;
            const args = arguments;
            clearTimeout(timeout);
            timeout = setTimeout(function () {
                func.apply(context, args);
            }, wait);
        };
    }

    // Rende la funzione debounce disponibile globalmente. Alcune parti dello script
    // richiedono l'accesso a debounce al di fuori del suo contesto originale.
    if (typeof window !== 'undefined') {
        window.debounce = debounce;
    }
      function exportData(type) {
        const data = loadData(type);
        const headers = ['Data/Ora','OV','OP','Codice','Descrizione','Lotto','Quantità','Stato'];
        const lines = [];
        lines.push(headers.join(';'));
        data.forEach(item => {
          const row = [ item.dateTime || '', item.ov || '', item.op || '', item.codice || '', item.descrizione || '', item.lotto || '', item.quantita || '', item.state || '' ].map(val => '"' + String(val).replace(/"/g, '""') + '"').join(';');
          lines.push(row);
        });
        const csv = lines.join('\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = `registro_sblocchi_${type}.csv`;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(a.href);
      }
      function exportBoth() {
        exportData('CQ');
        exportData('QA');
      }
      function printTables() {
        const tableCQ = document.getElementById('sbloccoTableCQ');
        const tableQA = document.getElementById('sbloccoTableQA');
        if (!tableCQ || !tableQA) return;
        // Apri una nuova finestra per la stampa.  Alcuni browser possono bloccare
        // le finestre popup restituendo null.  In tal caso, mostra un avviso
        // e interrompi la procedura di stampa per evitare errori JavaScript.
        const newWin = window.open('', '_blank');
        if (!newWin || !newWin.document) {
          // Se non è stato possibile aprire una nuova finestra, segnala
          // all'utente di abilitare i popup per completare la stampa.
          if (typeof alert === 'function') {
            alert('Impossibile aprire la finestra di stampa. Verifica le impostazioni del browser e consenti i popup per questa pagina.');
          }
          return;
        }
        newWin.document.write('<html><head><title>Registro Sblocchi CQ/QA</title>');
        newWin.document.write('<style>body { font-family: Arial, sans-serif; margin: 20px; } table { border-collapse: collapse; width: 100%; font-size: 12px; } th, td { border: 1px solid #ccc; padding: 4px; white-space: nowrap; } h3 { margin-top: 20px; }</style>');
        newWin.document.write('</head><body>');
        newWin.document.write('<h3>Sblocchi CQ</h3>');
        newWin.document.write(tableCQ.outerHTML);
        newWin.document.write('<h3>Sblocchi QA</h3>');
        newWin.document.write(tableQA.outerHTML);
        newWin.document.write('</body></html>');
        newWin.document.close();
        // Focalizza e stampa dopo un breve delay per garantire che il DOM sia
        // completamente caricato.  In seguito chiude la finestra di stampa.
        newWin.focus();
        setTimeout(function() {
          try {
            newWin.print();
          } catch (e) {}
          try {
            newWin.close();
          } catch (e) {}
        }, 500);
      }

      // Rendi disponibili globalmente alcune funzioni chiave in modo che il codice legacy
      // possa richiamarle (ad esempio per esportazione e stampa).  Senza queste
      // assegnazioni, exportBoth e printTables rimarrebbero chiuse nello scope
      // della funzione DOMContentLoaded e non sarebbero raggiungibili dal codice
      // esterno.
      window.exportBoth = exportBoth;
      window.printTables = printTables;
      // Espone anche i metodi di rendering per CQ e QA affinché possano essere
      // richiamati quando il modale viene aperto tramite funzioni legacy.
      window.renderCQ = function() { renderTable('CQ'); };
      window.renderQA = function() { renderTable('QA'); };
      // Espone le funzioni di caricamento e salvataggio in modo che possano essere
      // utilizzate al di fuori di questo scope (ad esempio dalla funzione
      // recordSbloccoEvent).  Senza queste assegnazioni non sarebbe possibile
      // accedere a loadData/saveData al di fuori di questa closure.
      window.sbloccoLoadData = loadData;
      window.sbloccoSaveData = saveData;
      function resetFilters(type) {
        const suffixes = ['StartDate','EndDate','StateFilter','FilterCodice','FilterOV','FilterOP','FilterDescrizione','FilterLotto'];
        suffixes.forEach(function(suffix) {
          const el = document.getElementById(`sblocco${type}${suffix}`);
          if (el) el.value = '';
        });
        renderTable(type);
      }
      function init() {
        const addCQ = document.getElementById('addCQRowBtn');
        if (addCQ) addCQ.addEventListener('click', function() { addRow('CQ'); });
        const addQA = document.getElementById('addQARowBtn');
        if (addQA) addQA.addEventListener('click', function() { addRow('QA'); });
        const resetCQ = document.getElementById('sbloccoCQResetBtn');
        if (resetCQ) resetCQ.addEventListener('click', function() { resetFilters('CQ'); });
        const resetQA = document.getElementById('sbloccoQAResetBtn');
        if (resetQA) resetQA.addEventListener('click', function() { resetFilters('QA'); });
        // Crea versioni debounce dei render per CQ e QA per ottimizzare i filtri
        const debouncedRenderCQ = debounce(() => renderTable('CQ'), 300);
        const debouncedRenderQA = debounce(() => renderTable('QA'), 300);
        ['StartDate','EndDate','StateFilter','FilterCodice','FilterOV','FilterOP','FilterDescrizione','FilterLotto'].forEach(function(suffix) {
          const elCQ = document.getElementById(`sbloccoCQ${suffix}`);
          if (elCQ) {
            const eventType = (suffix === 'StartDate' || suffix === 'EndDate' || suffix === 'StateFilter') ? 'change' : 'input';
            // Per gli eventi di input usa la funzione debounce per limitare i rendering ripetuti
            elCQ.addEventListener(eventType, function() {
              if (eventType === 'input') debouncedRenderCQ(); else renderTable('CQ');
            });
          }
          const elQA = document.getElementById(`sbloccoQA${suffix}`);
          if (elQA) {
            const eventType2 = (suffix === 'StartDate' || suffix === 'EndDate' || suffix === 'StateFilter') ? 'change' : 'input';
            elQA.addEventListener(eventType2, function() {
              if (eventType2 === 'input') debouncedRenderQA(); else renderTable('QA');
            });
          }
        });
        const exportBtn = document.getElementById('sbloccoExportBtn');
        if (exportBtn) exportBtn.addEventListener('click', function() { exportBoth(); });
        const printBtn = document.getElementById('sbloccoPrintBtn');
        if (printBtn) printBtn.addEventListener('click', function() { printTables(); });
        const openBtn = document.getElementById('sbloccoBtn');
        if (openBtn) {
          openBtn.addEventListener('click', function() {
            renderTable('CQ');
            renderTable('QA');
          });
        }
        renderTable('CQ');
        renderTable('QA');
      }
      init();
    });
    </script>
    <script>
        
// ===================================================================
    // ==> FUNZIONE DA AGGIUNGERE per il caricamento forzato dei dati <==
    // ===================================================================
    /**
     * Carica forzatamente dalla memoria del browser (localStorage) i dati
     * dei file statici (Layout, Referenze, Piano Analitico) per garantire
     * che siano sempre disponibili ad ogni avvio della pagina.
     */
    
// ===================================================================
    // ==> LOGICA PER IL NUOVO PULSANTE DI STAMPA GANTT MAGAZZINO <==
    // ===================================================================
    const printWarehouseGanttBtn = document.getElementById('printWarehouseGanttBtn');
    if (printWarehouseGanttBtn) {
        printWarehouseGanttBtn.addEventListener('click', () => {
            // Aggiunge la classe speciale al body per attivare gli stili CSS di stampa
            document.body.classList.add('printing-warehouse-gantt');

            // Imposta una funzione che verrà eseguita DOPO la stampa (sia che venga confermata o annullata)
            window.onafterprint = () => {
                // Rimuove la classe per ripristinare la visualizzazione normale
                document.body.classList.remove('printing-warehouse-gantt');
                // Pulisce l'evento per evitare che si attivi in altre stampe
                window.onafterprint = null; 
            };
            
            // Un piccolo ritardo per dare al browser il tempo di applicare gli stili
            setTimeout(() => {
                window.print(); // Apre la finestra di dialogo di stampa del browser
            }, 250);
        });
    }


document.getElementById('clearOpiFiltersBtn').addEventListener('click', function() {
    [
        'filterOpiOP','filterOpiOV','filterOpiCodice','filterOpiArticolo',
        'filterOpiCliente','filterOpiLotto','filterOpiQuantita','filterOpiUM','filterOpiOperatore'
    ].forEach(id => document.getElementById(id).value = '');
    applyOpiFilters();
});

document.addEventListener('DOMContentLoaded', () => {
flatpickr(document.getElementById('opiStartDate'), { dateFormat: "d/m/Y", locale: "it" });
flatpickr(document.getElementById('opiEndDate'), { dateFormat: "d/m/Y", locale: "it" });
flatpickr(document.getElementById('opiScadStartDate'), { dateFormat: "d/m/Y", locale: "it" });
flatpickr(document.getElementById('opiScadEndDate'), { dateFormat: "d/m/Y", locale: "it" });

 // ===================================================================
    // ==> CONFIGURAZIONE DEL SERVER <==
    // ===================================================================
    // Modifica questo indirizzo con l'IP del computer su cui gira XAMPP.
    // Per trovarlo: Apri Start, scrivi "cmd", nel terminale scrivi "ipconfig" e cerca "Indirizzo IPv4".
    const serverIP = '192.168.117'; // ESEMPIO: Sostituisci con l'IP reale del tuo server
    const apiEndpoint = `http://${serverIP}/programmazione/api.php`;

    const loginOverlay = document.getElementById('loginOverlay');
    const passwordInput = document.getElementById('passwordInput');
    const loginBtn = document.getElementById('loginBtn');
    const loginError = document.getElementById('loginError');
    const container = document.querySelector('.container');
    let currentUserLevel = 0;


    function caricaDatiStaticiForzatamente() {
        console.log("Eseguo caricamento forzato dei dati statici (Layout, Analisi)...");
        try {
            // Carica Layout
            const savedLayout = localStorage.getItem('layout_data');
            if (savedLayout) {
                layoutData = JSON.parse(savedLayout);
                // Non mostrare più il flag di stato del file layout.  Ora si utilizza solo la data di ultimo import.
                document.getElementById('layoutFileStatus').style.display = 'none';
                console.log("Dati Layout caricati da memoria.");
            }

            // Carica Referenze
            const savedReferenze = localStorage.getItem('referenzeData');
            if (savedReferenze) {
                referenzeData = JSON.parse(savedReferenze);
                // Non mostrare più il flag di stato del file referenze.  Ora si utilizza solo la data di ultimo import.
                document.getElementById('referenzeFileStatus').style.display = 'none';
                 console.log("Dati Referenze caricati da memoria.");
            }

            // Carica Piano Analitico
            const savedPianoAnalitico = localStorage.getItem('pianoAnaliticoData');
            if (savedPianoAnalitico) {
                pianoAnaliticoData = JSON.parse(savedPianoAnalitico);
                // Non mostrare più il flag di stato del piano analitico.  Ora si utilizza solo la data di ultimo import.
                document.getElementById('pianoAnaliticoFileStatus').style.display = 'none';
                 console.log("Dati Piano Analitico caricati da memoria.");
            }

            // IMPORTANTE: Esegue la logica di elaborazione solo se i dati necessari sono presenti
            if (referenzeData.length > 0 && pianoAnaliticoData.length > 0) {
                loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
                console.log("Dati di Analisi elaborati con successo.");
            }
        } catch (e) {
            console.error("Errore critico durante il caricamento forzato dei dati statici:", e);
            showAlert("Si è verificato un errore nel caricare i dati di configurazione salvati. Alcune funzionalità potrebbero non essere corrette.");
        }
    }




    const passwords = {
        "1234": 1,      // Utente Base
        "com123": 2,    // Utente Commerciale
        "pro234": 3,    // Utente Produzione
        "mag345": 4,    // Utente Magazzino
        "cq456": 5,     // Utente Analista
        "adm567": 6     // Amministratore
       };

    /*
     * Mappa dei permessi di allerta per ciascun livello utente.  Questi
     * valori derivano dal foglio Excel "Livelli.xlsx" e indicano per
     * ciascun ruolo se deve ricevere notifiche ADR, CQ o QA.  La chiave
     * corrisponde al livello definito in passwords; i valori sono booleani.
     */
    const alertPermissions = {
        // Livello 1: utente base.  Nessun permesso di notifica.
        1: { ADR: false, CQ: false, QA: false, quarantena: false, spedizioni: false },
        // Livello 2: utente commerciale.  Riceve tutte le notifiche CQ/QA e spedizioni.
        2: { ADR: true,  CQ: true,  QA: true,  quarantena: false, spedizioni: true  },
        // Livello 3: utente produzione.  Riceve tutte le notifiche CQ/QA e spedizioni.
        3: { ADR: true,  CQ: true,  QA: true,  quarantena: false, spedizioni: true  },
        // Livello 4: utente magazzino.  Riceve tutte le notifiche CQ/QA e spedizioni.
        4: { ADR: true,  CQ: true,  QA: true,  quarantena: false, spedizioni: true  },
        // Livello 5 (Analista) corrisponde al Controllo Qualità.  Riceve solo la quarantena.
        5: { ADR: true,  CQ: false, QA: false, quarantena: true, spedizioni: false  },
        // Livello 6 (Amministratore).  Riceve solo ADR.
        6: { ADR: true,  CQ: false, QA: false, quarantena: false, spedizioni: false }
    };

    /*
     * Mappa dei permessi per la generazione della Packing List.  La chiave
     * corrisponde al livello utente definito in passwords; il valore
     * indica se l'utente può creare una Packing List.  Questi valori
     * sono derivati dal file "Livelli.xlsx":
     * - Livello 1 (Base):            no
     * - Livello 2 (Commerciale):     si
     * - Livello 3 (Produzione):      si
     * - Livello 4 (Magazzino):       si
     * - Livello 5 (Analista):        no
     * - Livello 6 (Amministratore):  si
     */
    const packingListPermissions = {
        1: false,
        2: true,
        3: true,
        4: true,
        5: false,
        6: true
    };

     function applyPermissions(level) {
    console.log('Applico permessi per livello:', level);
    const allInputs = document.querySelectorAll('input:not(#passwordInput), select, textarea');
    const allButtons = document.querySelectorAll('button:not(#loginBtn)');

    // 1. Blocco totale di default
    allInputs.forEach(el => {
        el.readOnly = true;
        el.disabled = true;
        el.style.pointerEvents = 'none';
    });
    allButtons.forEach(btn => btn.style.display = 'none');

    // 2. Abilitazioni di base per tutti (filtri, esportazioni, ecc.)
    const baseControls = document.querySelectorAll(
        '#searchInput, #findBtn, #findNextBtn, #clearFilterBtn, #filterColumn1, #filterValue1, #filterColumn2, #filterValue2, ' +
        '#scrollLeftBtn, #scrollRightBtn, #exportDataBtn, #exportDailyPdfBtn, #exportDailyWordBtn, #exportDailyExcelBtn, ' +
        '#filterDailyColumn, #filterDailyValue, #clearDailyFilterBtn, #exportShippingDataBtn, #printShippingBtn, ' +
        '#filterShippingColumn, #filterShippingValue, #clearShippingFilterBtn, #clearShippingDateBtn, ' +
        '#exportArrivalDataBtn, #exportPropostaLayoutBtn, #filterArrivalColumn, #filterArrivalValue, ' +
        '#clearArrivalFilterBtn, #clearArrivalDateBtn, #loadDataBtn, #saveDataBtn, #exportAnalisiPdfBtn, #manualRefreshBtn, ' +
        '#medicalDeviceStartDate, #medicalDeviceEndDate, #clearMedicalDeviceDateBtn, ' +
        '#filterMedicalDeviceCodice, #filterMedicalDeviceDescrizione, #filterMedicalDeviceCliente, #clearMedicalDeviceFiltersBtn'
    );
    baseControls.forEach(el => {
        if(el) {
            el.style.display = 'inline-flex';
            el.disabled = false;
            el.style.pointerEvents = 'auto';
        }
    });

    // 3. Sblocco selettivo basato sul ruolo
    switch (level) {
        case 2: // Commerciale
            document.querySelectorAll('#addSalesOrderRowBtn, #duplicateSalesOrderRowBtn, #deleteSalesOrderRowBtn, #importOVBtn, #sendEmailOVBtn').forEach(btn => { if(btn){ btn.style.display = 'inline-flex'; btn.disabled = false; }});
            document.querySelectorAll('#salesOrderTable input, #salesOrderTable select').forEach(el => { el.readOnly = false; el.disabled = false; el.style.pointerEvents = 'auto'; });

            // Dopo aver applicato i permessi al commerciale, verifica e notifica
            // eventuali spedizioni ADR per questa settimana.  La notifica è
            // indipendente per ciascun ruolo (commerciale e magazzino).
            if (typeof checkAndNotifyADR === 'function') {
                try {
                    checkAndNotifyADR();
                } catch (e) {
                    console.warn('Errore nel controllo ADR dopo login commerciale:', e);
                }
            }

                // Assicurati che i pulsanti del pop‑up ADR siano visibili e interattivi
                const adrPostponeBtn = document.getElementById('adrPostponeBtn');
                const adrAckBtn = document.getElementById('adrAcknowledgeBtn');
                if (adrPostponeBtn && adrAckBtn) {
                    adrPostponeBtn.style.display = 'inline-flex';
                    adrAckBtn.style.display = 'inline-flex';
                    adrPostponeBtn.disabled = false;
                    adrAckBtn.disabled = false;
                    adrPostponeBtn.style.pointerEvents = 'auto';
                    adrAckBtn.style.pointerEvents = 'auto';
                }

            // Il pulsante Packing List verrà gestito globalmente in base alla mappa dei permessi.

            // Per il ruolo commerciale rendi visibile anche il pulsante di importazione del
            // programma di produzione (Importa PP).  Secondo la tabella dei permessi
            // fornita dall’utente (Livelli.xlsx) l'utente commerciale deve poter
            // effettuare l'import del piano di produzione.  Attiviamo quindi il
            // pulsante e ne abilitiamo gli eventi del puntatore.
            {
                const importPP = document.getElementById('importPPBtn');
                if (importPP) {
                    importPP.style.display = 'inline-flex';
                    importPP.disabled = false;
                    importPP.style.pointerEvents = 'auto';
                }
            }
            break;
        case 3: // Produzione
            // Abilita i pulsanti principali e garantisce che gli eventi del puntatore siano attivi
            document.querySelectorAll('#addRowBtn, #duplicateRowBtn, #deleteRowBtn, #importPPBtn, #sendEmailBtn').forEach(btn => {
                if(btn){
                    btn.style.display = 'inline-flex';
                    btn.disabled = false;
                    btn.style.pointerEvents = 'auto';
                }
            });
            document.querySelectorAll('#productionTable input, #productionTable select, #productionTable textarea').forEach(el => { el.readOnly = false; el.disabled = false; el.style.pointerEvents = 'auto'; });
            document.querySelectorAll('#addDailyRowBtn, #duplicateDailyRowBtn, #deleteDailyRowBtn, #saveDailyDataBtn, #loadDailyDataBtn').forEach(btn => { if(btn){ btn.style.display = 'inline-flex'; btn.disabled = false; }});
            document.querySelectorAll('#dailyProductionTable input, #dailyProductionTable select').forEach(el => { el.readOnly = false; el.disabled = false; el.style.pointerEvents = 'auto'; });
            document.querySelectorAll('#addMedicalDeviceRowBtn').forEach(btn => { if(btn){ btn.style.display = 'inline-flex'; btn.disabled = false; }});
            document.querySelectorAll('#medicalDeviceProductionTable .scarti-input').forEach(el => { el.readOnly = false; el.disabled = false; el.style.pointerEvents = 'auto'; });

            // L'utente produzione non deve vedere l'avviso ADR.  Nascondi il pop-up se presente.
            {
                const adrDivProd = document.getElementById('adrNotification');
                if (adrDivProd) {
                    adrDivProd.style.display = 'none';
                }
            }

            // Rende visibile il pulsante Packing List per il ruolo produzione (analogamente alla versione 66)
            {
                const pbtn = document.getElementById('packingListBtn');
                if (pbtn) {
                    pbtn.style.display = 'inline-flex';
                    pbtn.disabled = false;
                    pbtn.style.pointerEvents = 'auto';
                }
            }
            // Rendi visibile anche il pulsante Sblocchi CQ/QA e i suoi controlli nel ruolo Produzione.
            {
                const sbBtn = document.getElementById('sbloccoBtn');
                if (sbBtn) {
                    sbBtn.style.display = 'inline-flex';
                    sbBtn.disabled = false;
                    sbBtn.style.pointerEvents = 'auto';
                }
                // Anche i pulsanti del modale (Esporta, Stampa, Chiudi) devono essere interattivi
                const sbExport = document.getElementById('sbloccoExportBtn');
                const sbPrint = document.getElementById('sbloccoPrintBtn');
                const sbClose = document.getElementById('sbloccoCloseBtn');
                [sbExport, sbPrint, sbClose].forEach(btn => {
                    if (btn) {
                        btn.style.display = 'inline-flex';
                        btn.disabled = false;
                        btn.style.pointerEvents = 'auto';
                    }
                });

                // Rendi visibili i pulsanti interni al registro sblocchi (aggiungi riga e reset) per CQ e QA
                [
                  document.getElementById('addCQRowBtn'),
                  document.getElementById('addQARowBtn'),
                  document.getElementById('sbloccoCQResetBtn'),
                  document.getElementById('sbloccoQAResetBtn')
                ].forEach(btn => {
                  if (btn) {
                    btn.style.display = 'inline-flex';
                    btn.disabled = false;
                    btn.style.pointerEvents = 'auto';
                  }
                });
            }
            break;
        case 4: // Magazzino
    // Pulsanti principali di magazzino
    document.querySelectorAll(
        '#addShippingRowBtn, #duplicateShippingRowBtn, #deleteShippingRowBtn, #importOSBtn, #sendShippingEmailBtn, #importOpiBtn, #sendOpiEmailBtn'
    ).forEach(btn => {
        if (btn) {
            btn.style.display = 'inline-flex';
            btn.disabled = false;
            // Riabilita esplicitamente gli eventi del puntatore per questi pulsanti.
            btn.style.pointerEvents = 'auto';
        }
    });

    // Tabelle spedizioni
    document.querySelectorAll('#shippingScheduleTable input, #shippingScheduleTable select').forEach(el => {
        el.readOnly = false; el.disabled = false; el.style.pointerEvents = 'auto';
    });

    // Tabelle arrivi
    document.querySelectorAll('#addArrivalRowBtn, #duplicateArrivalRowBtn, #deleteArrivalRowBtn, #importArrivalsBtn, #importLayoutBtn, #sendArrivalEmailBtn, #importOVBtn').forEach(btn => {
        if (btn) {
            btn.style.display = 'inline-flex';
            btn.disabled = false;
            // Riabilita esplicitamente gli eventi del puntatore per i pulsanti legati agli arrivi.
            btn.style.pointerEvents = 'auto';
        }
    });
    document.querySelectorAll('#arrivalScheduleTable input, #arrivalScheduleTable select').forEach(el => {
        el.readOnly = false; el.disabled = false; el.style.pointerEvents = 'auto';
    });

    // Tutti i filtri e campi di magazzino (incl. OPI e Merce non arrivata)
    [
        // OPI
        'opiStartDate','opiEndDate','opiScadStartDate','opiScadEndDate',
        'filterOpiOP','filterOpiOV','filterOpiCodice','filterOpiArticolo',
        'filterOpiCliente','filterOpiLotto','filterOpiQuantita','filterOpiUM','filterOpiOperatore',
        'clearOpiFiltersBtn','importOpiBtn','sendOpiEmailBtn',

        // Spedizioni
        'filterShippingColumn','filterShippingValue','clearShippingFilterBtn','clearShippingDateBtn',
        'shippingStartDate','shippingEndDate',

        // Arrivi
        'filterArrivalColumn','filterArrivalValue','clearArrivalFilterBtn','clearArrivalDateBtn',
        'arrivalStartDate','arrivalEndDate',

        // Merce non arrivata (Not Arrived)
        'filterOverdueOV','filterOverdueCodice','filterOverdueDescrizione','filterOverdueRagSoc',
        'filterOverdueDataDa','filterOverdueDataA','clearOverdueFiltersBtn'
    ].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.style.display = 'inline-flex';
            el.disabled = false;
            el.readOnly = false;
            el.style.pointerEvents = 'auto';
        }
    });

    // Tabella Merce Non Arrivata (lettura campi filtro)
    document.querySelectorAll('#overdueArrivalsTable input, #overdueArrivalsTable select').forEach(el => {
        el.readOnly = true; // tabella in sola lettura
        el.disabled = true; // ma i filtri rimangono abilitati!
    });

            // Dopo aver applicato i permessi al magazzino, verifica e notifica
            // eventuali spedizioni ADR per questa settimana.  La notifica è
            // indipendente per ciascun ruolo (commerciale e magazzino).
            if (typeof checkAndNotifyADR === 'function') {
                try {
                    checkAndNotifyADR();
                } catch (e) {
                    console.warn('Errore nel controllo ADR dopo login magazzino:', e);
                }
            }

            // Assicurati che i pulsanti del pop‑up ADR siano visibili e interattivi
            const adrPostponeBtn4 = document.getElementById('adrPostponeBtn');
            const adrAckBtn4 = document.getElementById('adrAcknowledgeBtn');
            if (adrPostponeBtn4 && adrAckBtn4) {
                adrPostponeBtn4.style.display = 'inline-flex';
                adrAckBtn4.style.display = 'inline-flex';
                adrPostponeBtn4.disabled = false;
                adrAckBtn4.disabled = false;
                adrPostponeBtn4.style.pointerEvents = 'auto';
                adrAckBtn4.style.pointerEvents = 'auto';
            }
            // Il pulsante Packing List verrà gestito globalmente in base alla mappa dei permessi.
    break;
        case 5: // Analista CQ
            document.querySelectorAll('#importReferenzeBtn, #importPianoAnaliticoBtn, #importDeviceRefBtn, #importMedicalProductionBtn, #addAnalisiRowBtn, #duplicateAnalisiRowBtn, #deleteAnalisiRowBtn').forEach(btn => { if(btn){ btn.style.display = 'inline-flex'; btn.disabled = false; }});
            document.querySelectorAll('#analisiTable input, #analisiTable select, #analisiTable textarea, #searchLottoInput').forEach(el => { el.readOnly = false; el.disabled = false; el.style.pointerEvents = 'auto'; });

            // L'analista CQ non deve vedere l'avviso ADR.  Se esiste un pop-up,
            // nascondilo esplicitamente.
            {
                const adrDivCQ = document.getElementById('adrNotification');
                if (adrDivCQ) {
                    adrDivCQ.style.display = 'none';
                }
            }
            break;
        case 6: // Amministratore
            allInputs.forEach(el => { el.disabled = false; el.readOnly = false; el.style.pointerEvents = 'auto'; });
            allButtons.forEach(btn => { if(btn){ btn.style.display = 'inline-flex'; btn.disabled = false;} });
            // L'amministratore non deve vedere l'avviso ADR.  Nascondi il pop-up se presente.
            const adrDivAdmin = document.getElementById('adrNotification');
            if (adrDivAdmin) {
                adrDivAdmin.style.display = 'none';
            }

            // Il pulsante Packing List verrà gestito globalmente in base alla mappa dei permessi.
            break;
    }
    
    // Gestisce la visibilità del pulsante Packing List in base ai permessi definiti nella mappa.
    {
        const pbtn = document.getElementById('packingListBtn');
        if (pbtn) {
            if (packingListPermissions[level]) {
                pbtn.style.display = 'inline-flex';
                pbtn.disabled = false;
                pbtn.style.pointerEvents = 'auto';
            } else {
                // Nasconde completamente il pulsante se l'utente non ha il permesso di generare la Packing List
                pbtn.style.display = 'none';
            }
        }
    }

    container.style.visibility = 'visible';
    // Dopo aver applicato i permessi, verifica se ci sono avvisi CQ/QA da mostrare
    if (typeof checkAndNotifyQuality === 'function') {
        try {
            checkAndNotifyQuality();
        } catch (e) {
            console.warn('Errore nel controllo degli avvisi CQ/QA:', e);
        }
    }
}

   function handleLogin() {
    console.log('Login livello:', currentUserLevel);
    const enteredPassword = passwordInput.value;
    currentUserLevel = passwords[enteredPassword] || 0;

    if (currentUserLevel > 0) {
        loginOverlay.style.display = 'none';
        applyPermissions(currentUserLevel);

        // Ricarica la tabella analisi dopo il login per applicare i permessi corretti
        console.log(`Login successo con livello: ${currentUserLevel}. Ricarico la tabella analisi.`);
        updateAnalisiTable();

        // Esegui l'inizializzazione completa dell'applicazione dopo il login.
        // Questo carica i dati salvati e genera i grafici solo quando l'utente è autentificato,
        // così da evitare rallentamenti all'apertura della pagina.
        initializeAfterLogin();

        // Salva il livello dell'utente per ripristinare il login dopo un refresh
        try {
            sessionStorage.setItem('userLevel', String(currentUserLevel));
        } catch (e) {
            console.warn('Impossibile salvare il livello utente in sessionStorage:', e);
        }
        
    } else {
        loginError.style.display = 'block';
        passwordInput.value = '';
    }
}

    loginBtn.addEventListener('click', handleLogin);
    passwordInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            handleLogin();
        }
    });

    container.style.visibility = 'hidden';

    const productionTableBody = document.querySelector('#productionTable tbody');
    const addRowBtn = document.getElementById('addRowBtn');
    const duplicateRowBtn = document.getElementById('duplicateRowBtn');
    const deleteRowBtn = document.getElementById('deleteRowBtn');
    const saveDataBtn = document.getElementById('saveDataBtn');
    const loadDataBtn = document.getElementById('loadDataBtn');
    const manualRefreshBtn = document.getElementById('manualRefreshBtn');
    const sendEmailBtn = document.getElementById('sendEmailBtn');
    const currentDateSpan = document.getElementById('currentDate');
    const currentWeekSpan = document.getElementById('currentWeek');
    const importPPBtn = document.getElementById('importPPBtn');
    const exportDataBtn = document.getElementById('exportDataBtn');
    const fileInput = document.getElementById('fileInput');
    const overdueArrivalsTableBody = document.querySelector('#overdueArrivalsTable tbody');
    const ganttChartDiv = document.getElementById('ganttChart');
    const warehouseGanttChartDiv = document.getElementById('warehouseGanttChart');
    let genericTooltip = document.getElementById('genericTooltip');
    const tableContainer = document.querySelector('.table-container');
    const scrollLeftBtn = document.getElementById('scrollLeftBtn');
    const scrollRightBtn = document.getElementById('scrollRightBtn');
    const customModal = document.getElementById('customModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalMessage = document.getElementById('modalMessage');
    const modalButtons = document.getElementById('modalButtons');
    const searchInput = document.getElementById('searchInput');
    const findBtn = document.getElementById('findBtn');
    const findNextBtn = document.getElementById('findNextBtn');
    let searchResults = [];
    let currentSearchIndex = -1;
    const filterColumn1Select = document.getElementById('filterColumn1');
    const filterValue1Input = document.getElementById('filterValue1');
    let flatpickrInstance1 = null;
    const filterColumn2Select = document.getElementById('filterColumn2');
    const filterValue2Input = document.getElementById('filterValue2');
    let flatpickrInstance2 = null;
    const dateFilterColumns = ['produzioneData', 'dataConfezionamento'];
    const medicalDevicesFilterValue = 'medicalDevices';
    let importMode = null;

    /*
     * Mappa dei codici di confezionamento sterili.
     * Questa mappa è generata dal file "Codici Sterili.xls" fornito
     * dall’utente.  Ogni chiave rappresenta il codice prodotto (colonna B)
     * e il valore associato è il corrispondente codice di confezionamento
     * (colonna C) da inserire automaticamente nella colonna "Codice
     * Confezionamento" quando viene creato o importato un record di
     * produzione.  Se un codice prodotto non è presente in questa
     * mappatura, verrà applicata la logica predefinita (code*-variant
     * oppure code-KG per pezzo).
     */
    const sterilePackagingMap = {
      "6267": "6267*",
      "6268": "6268*",
      "6269": "6269*",
      "6271": "6271*",
      "6274": "6274-A",
      "6275": "6275*",
      "6276": "6276-A",
      "6277": "6277*",
      "6278": "6278*",
      "6279": "6279*",
      "6280": "6280*",
      "6282": "6282*",
      "6283": "6283*",
      "6878": "6878*",
      "6879": "6879*",
      "6955": "6955*",
      "7299": "7299*",
      "7383": "7383*",
      "7384": "7384*",
      "7489": "7489*",
      "7561": "7561*",
      "7567": "7567*",
      "7569": "7569*",
      "7594": "7594*",
      "7595": "7595*",
      "7602": "7602*",
      "7603": "7603*",
      "7658": "7658",
      "7685": "7685-A",
      "7694": "7694*",
      "7699": "7699-A",
      "7731": "7731*",
      "7743": "7743*",
      "7744": "7744-0,005",
      "7745": "7745*",
      "7746": "7746*",
      "7748": "7748*",
      "7770": "7770-10F-I/R",
      "7771": "7771-10F-I/R",
      "7781": "7781*",
      "7806": "7806*",
      "7807": "7807*",
      "7901": "7901*",
      "7909": "7909*",
      "7911": "7911*",
      "7938": "7938*-I/R",
      "7974": "7974*",
      "7975": "7975*",
      "7978": "7978*"
    };
    const stickyControlsWrapper = document.getElementById('stickyControlsWrapper');
    const tableHead = document.querySelector('#productionTable thead');
    const unitOptions = ['Kg', 'Pz', 'mL', 'L', 'g'];
    const dailyProductionDateInput = document.getElementById('dailyProductionDateInput');
    const dailyProductionTableBody = document.querySelector('#dailyProductionTable tbody');
    const filterDailyColumnSelect = document.getElementById('filterDailyColumn');
    const filterDailyValueInput = document.getElementById('filterDailyValue');
    const applyDailyFilterBtn = document.getElementById('applyDailyFilterBtn');
    const clearDailyFilterBtn = document.getElementById('clearDailyFilterBtn');
    const exportDailyPdfBtn = document.getElementById('exportDailyPdfBtn');
    let dailyProductionFlatpickr = null;
    let dailyProductionSelectedDate = null;
    let dailyProductionOperatorFilter = '';

    // Gestione dinamica della posizione dei pulsanti di scorrimento orizzontale per la tabella di produzione.
    // I pulsanti devono restare visibili quando l'utente scorre verticalmente la pagina.  Calcoliamo
    // la parte visibile della tableContainer e centriamo verticalmente il wrapper.  Se la tabella
    // non è visibile li nascondiamo.
    const scrollButtonsWrapper = document.querySelector('.scroll-buttons-wrapper');
    function repositionScrollButtons() {
        if (!tableContainer || !scrollButtonsWrapper) return;
        const rect = tableContainer.getBoundingClientRect();
        // Nascondi se la tabella è fuori vista
        if (rect.bottom < 0 || rect.top > window.innerHeight) {
            scrollButtonsWrapper.style.display = 'none';
            return;
        }
        scrollButtonsWrapper.style.display = 'flex';
        const visibleTop = Math.max(rect.top, 0);
        const visibleBottom = Math.min(rect.bottom, window.innerHeight);
        const visibleHeight = visibleBottom - visibleTop;
        const wrapperHeight = scrollButtonsWrapper.offsetHeight || 0;
        const top = visibleTop + visibleHeight / 2 - wrapperHeight / 2;
        scrollButtonsWrapper.style.top = top + 'px';
    }
    if (scrollButtonsWrapper) {
        if (tableContainer) tableContainer.addEventListener('scroll', repositionScrollButtons);
        window.addEventListener('scroll', repositionScrollButtons);
        window.addEventListener('resize', repositionScrollButtons);
        // Inizializza la posizione quando il layout è pronto
        setTimeout(repositionScrollButtons, 0);
    }

    // Datalist per suggerimenti operatori e macchinari nel programma giornaliero
    const operatorSuggestionsList = document.getElementById('operatorSuggestionsList');
    const macchinariOptionsListDaily = document.getElementById('macchinariOptionsListDaily');
    const dailyOperationsOptions = [
        "",
        "Check materie prime - lavorazione - insiringamento - autoclave",
        "Inflaconare\n- etichettare\n- lottare\n- astucciare",
        "Sperlatura ed etichettatura siringhe",
        "Inserimento e/o etichettatura blister",
        "Costruzione e/o etichettatura scatola - inscatolamento",
        "Controllo prodotto - confezionamento - conformità OLS",
        "Filtrare - campionamento",
        "Check materie prime - lavorazione - campionamento",
        "Campo libero"
    ];
    const salesOrderTableBody = document.querySelector('#salesOrderTable tbody');
    const addSalesOrderRowBtn = document.getElementById('addSalesOrderRowBtn');
    const duplicateSalesOrderRowBtn = document.getElementById('duplicateSalesOrderRowBtn');
    const deleteSalesOrderRowBtn = document.getElementById('deleteSalesOrderRowBtn');
    const importOVBtn = document.getElementById('importOVBtn');
    const sendEmailOVBtn = document.getElementById('sendEmailOVBtn');
    const logbookContentElement = document.getElementById('logbookContent');
    const logbookContainer = document.getElementById('logbookContainer');
    const logbookBtn = document.getElementById('logbookBtn');
    const clearLogbookBtn = document.getElementById('clearLogbookBtn');
    const printLogbookBtn = document.getElementById('printLogbookBtn');
    const logbookStartDateInput = document.getElementById('logbookStartDate');
    const logbookEndDateInput = document.getElementById('logbookEndDate');
    const logbookStartTime = document.getElementById('logbookStartTime');
    const logbookEndTime = document.getElementById('logbookEndTime');
    const clearLogbookFilterBtn = document.getElementById('clearLogbookFilterBtn');
    let logbookEntries = [];
    const importReferenzeBtn = document.getElementById('importReferenzeBtn');
    const importPianoAnaliticoBtn = document.getElementById('importPianoAnaliticoBtn');
    const referenzeInput = document.getElementById('referenzeInput');
    const pianoAnaliticoInput = document.getElementById('pianoAnaliticoInput');
    const analisiTableBody = document.querySelector('#analisiTable tbody');
    const analisiTableHeaders = document.getElementById('analisiHeaders');
    const referenzeFileStatusSpan = document.getElementById('referenzeFileStatus');
    const pianoAnaliticoFileStatusSpan = document.getElementById('pianoAnaliticoFileStatus');
    const addAnalisiRowBtn = document.getElementById('addAnalisiRowBtn');
    const duplicateAnalisiRowBtn = document.getElementById('duplicateAnalisiRowBtn');
    const clearAnalisiDateBtn = document.getElementById('clearAnalisiDateBtn');
    const exportAnalisiPdfBtn = document.getElementById('exportAnalisiPdfBtn');
    const importOSBtn = document.getElementById('importOSBtn');
    const shippingScheduleTableBody = document.querySelector('#shippingScheduleTable tbody');
    const importOpiBtn = document.getElementById('importOpiBtn');
// Nuovo bottone per l'import dei riferimenti dispositivi (DeviceRef)
const importDeviceRefBtn = document.getElementById('importDeviceRefBtn');
    const opiTableBody = document.querySelector('#opiTable tbody');
// ========================================================================
    // ==> NUOVE VARIABILI PER LA TABELLA "PROGRAMMA GIORNALIERO DI ARRIVO MERCE"
    // ========================================================================
    const importArrivalsBtn = document.getElementById('importArrivalsBtn');
    const arrivalScheduleTableBody = document.querySelector('#arrivalScheduleTable tbody');
    const addArrivalRowBtn = document.getElementById('addArrivalRowBtn');
    const duplicateArrivalRowBtn = document.getElementById('duplicateArrivalRowBtn');
    const deleteArrivalRowBtn = document.getElementById('deleteArrivalRowBtn');
    const sendArrivalEmailBtn = document.getElementById('sendArrivalEmailBtn');
    const exportArrivalDataBtn = document.getElementById('exportArrivalDataBtn');
    const arrivalStartDateInput = document.getElementById('arrivalStartDate');
    const arrivalEndDateInput = document.getElementById('arrivalEndDate');
    const filterArrivalColumn = document.getElementById('filterArrivalColumn');
    const filterArrivalValue = document.getElementById('filterArrivalValue');
    const applyArrivalFilterBtn = document.getElementById('applyArrivalFilterBtn');
    const clearArrivalFilterBtn = document.getElementById('clearArrivalFilterBtn');
    const clearArrivalDateBtn = document.getElementById('clearArrivalDateBtn');
    const addShippingRowBtn = document.getElementById('addShippingRowBtn');
    const duplicateShippingRowBtn = document.getElementById('duplicateShippingRowBtn');
    const deleteShippingRowBtn = document.getElementById('deleteShippingRowBtn');
    const sendShippingEmailBtn = document.getElementById('sendShippingEmailBtn');
    const exportShippingDataBtn = document.getElementById('exportShippingDataBtn');
    const shippingScheduleDateInput = document.getElementById('shippingScheduleDateInput');

    // ========================================================================
    // ==> NUOVE VARIABILI PER LA TABELLA "MERCE IN QUARANTENA" E STATO MAGAZZINO
    // ========================================================================
    // Corpo della tabella quarantena: serve per aggiungere o rimuovere righe
    const quarantineTableBody = document.querySelector('#quarantineTable tbody');
    // Flag che indica se la password magazzino è stata già verificata in questa sessione
    let magPasswordValidated = false;
    const shippingStartDateInput = document.getElementById('shippingStartDate');
    const shippingEndDateInput = document.getElementById('shippingEndDate');
    const printShippingBtn = document.getElementById('printShippingBtn');
    if (printShippingBtn) {
        printShippingBtn.addEventListener('click', () => {
            // Applica la classe per isolare la sezione durante la stampa
            document.body.classList.add('printing-shipping');
            
            // Imposta una funzione per rimuovere la classe dopo la stampa
            window.onafterprint = () => {
                document.body.classList.remove('printing-shipping');
                window.onafterprint = null; // Pulisce l'handler
            };

            // Avvia la stampa
            window.print();

// --- PATCH: ensure Warehouse Gantt controls are visible/enabled ---
try {
  [
    '#warehouseGanttScrollLeftBtn',
    '#warehouseGanttScrollRightBtn',
    '#warehouseGanttScrollButtonsWrapper',
    '#warehouseGanttExternalScrollbar'
  ].forEach(function(sel){
    var el = document.querySelector(sel);
    if (el) {
      el.style.display = (sel === '#warehouseGanttExternalScrollbar') ? 'block' : 'inline-flex';
      el.style.pointerEvents = 'auto';
      if ('disabled' in el) el.disabled = false;
    }
  });
} catch (e) {
  console.warn('Patch: could not toggle Warehouse Gantt controls', e);
}
// --- END PATCH ---

// --- PATCH: show/enable Warehouse Gantt top scrollbar + side buttons ---
try {
  ['#warehouseGanttExternalScrollbar','#warehouseGanttScrollButtonsWrapper','#warehouseGanttScrollLeftBtn','#warehouseGanttScrollRightBtn'].forEach(function(sel){
    var el = document.querySelector(sel);
    if (el) {
      el.style.display = sel === '#warehouseGanttExternalScrollbar' ? 'block' : 'inline-flex';
      el.style.pointerEvents = 'auto';
      if ('disabled' in el) el.disabled = false;
    }
  });
} catch(e){ console.warn('applyPermissions patch: Gantt controls', e); }
// --- END PATCH ---
});
    }

    // ================================================================
    // ==> CARICAMENTO INIZIALE DELLA TABELLA QUARANTENA
    // ================================================================
    // Al caricamento della pagina prova a popolare la tabella "Merce in Quarantena"
    // con i dati già presenti nel localStorage.  In questo modo le righe evase
    // restano visibili anche prima che arrivi una risposta dal server.  Se non
    // vi sono dati memorizzati o la struttura è corrotta, la tabella rimane vuota.
    try {
        const storedQuarantine = localStorage.getItem('quarantine_data_autosave');
        if (storedQuarantine) {
            const parsedQuarantine = JSON.parse(storedQuarantine);
            if (Array.isArray(parsedQuarantine) && typeof populateQuarantineTable === 'function') {
                populateQuarantineTable(parsedQuarantine);
            }
        }
    } catch (e) {
        console.warn('Impossibile caricare i dati della quarantena dal localStorage:', e);
    }

    let referenzeData = [];
    let pianoAnaliticoData = [];

    
    function saveStaticData(key, data, filename) {
        try {
            localStorage.setItem(key, JSON.stringify(data));
            localStorage.setItem(`${key.replace('Data', 'FileName')}`, filename);
            console.log(`Dati per ${key} salvati in localStorage.`);
        } catch (e) {
            console.error(`Errore nel salvataggio dei dati per ${key} in localStorage:`, e);
        }
    }

    function loadAnalisiExcelData(pianoAnaliticoCsv, referenzeCsv) {
        // Se manca uno dei due file, non procedere
        if (!pianoAnaliticoCsv || !referenzeCsv) {
            console.warn("Impossibile caricare i dati di analisi: mancano i file di riferimento.");
            return;
        }
        const mdPianoAnaliticoRows = pianoAnaliticoCsv;
        const headerRowIndex = 4;
        const methodsRowIndex = 5;
        const dataStartIndex = 6;
        const dataColOffset = 4;

        analysisHeaders = mdPianoAnaliticoRows[headerRowIndex].slice(dataColOffset);
        methodHeaders = mdPianoAnaliticoRows[methodsRowIndex].slice(dataColOffset);

        analysisPlan = {};
        for (let i = dataStartIndex; i < mdPianoAnaliticoRows.length; i++) {
            const row = mdPianoAnaliticoRows[i];
            if (row.length < 2) continue;
            const productCode = String(row[1]).trim();
            const productName = String(row[2]).trim();

            if (!productCode) continue;

            analysisPlan[productCode] = { name: productName, analyses: {} };
            for (let j = 0; j < analysisHeaders.length; j++) {
                const analysisName = analysisHeaders[j];
                const cellContent = String(row[j + dataColOffset] || '').trim();
                if (cellContent !== '' && cellContent.toUpperCase() !== 'N.A.' && cellContent.toUpperCase() !== 'NA') {
                    analysisPlan[productCode].analyses[analysisName] = true;
                }
            }
        }

        referenzeMap = {};
        for (let i = 2; i < referenzeCsv.length; i++) {
            const row = referenzeCsv[i];
            if (row.length < 5) continue;
            const prodottoCapostipiteRef = String(row[1]).trim();
            const semiconfezionatoRef = String(row[2]).trim();
            const referenzaEquivalente = String(row[3]).trim();
            const referenzaEquivalenteRef = String(row[4]).trim();

            if (!referenzeMap[referenzaEquivalenteRef]) {
                referenzeMap[referenzaEquivalenteRef] = {};
            }
            if (!referenzeMap[referenzaEquivalenteRef].codiceFinito) {
                referenzeMap[referenzaEquivalenteRef].codiceFinito = prodottoCapostipiteRef;
            }
            if (!referenzeMap[referenzaEquivalenteRef].codiceSemi) {
                referenzeMap[referenzaEquivalenteRef].codiceSemi = semiconfezionatoRef;
            }
            if (!referenzeMap[referenzaEquivalenteRef].codiceBulk) {
                // il codice bulk per i semi confezionati è derivato rimuovendo l'asterisco *SC
                referenzeMap[referenzaEquivalenteRef].codiceBulk = semiconfezionatoRef.replace(/\*SC$/, '');
            }
            if (!referenzeMap[referenzaEquivalenteRef].nomeProdotto) {
                referenzeMap[referenzaEquivalenteRef].nomeProdotto = referenzaEquivalente;
            }
        }
    }

    function deriveBulkLotto(semiLotto) {
        if (!semiLotto || typeof semiLotto !== 'string' || semiLotto.length < 4) { // Minimo 4 cifre es. 00125
            return semiLotto;
        }
        const yearSuffix = semiLotto.slice(-2);
        const prefix = semiLotto.slice(0, -2);

        const prefixNum = parseInt(prefix, 10);
        if (isNaN(prefixNum)) {
            return semiLotto;
        }

        const decrementedPrefix = prefixNum - 1;
        // Ricostruisce il prefisso con gli zeri iniziali
        const newPrefix = String(decrementedPrefix).padStart(prefix.length, '0');

        return `${newPrefix}${yearSuffix}`;
    }

    function updateLastModifiedTimestamp() {
        const now = new Date();
        const formattedTimestamp = now.toLocaleString('it-IT', {
            day: '2-digit', month: '2-digit', year: 'numeric',
            hour: '2-digit', minute: '2-digit'
        });
        const timestampSpan = document.getElementById('lastModifiedTimestamp');
        if (timestampSpan) {
            timestampSpan.textContent = formattedTimestamp;
        }
        localStorage.setItem('last_modified_timestamp', formattedTimestamp);
    }

    function loadLastModifiedTimestamp() {
        const savedTimestamp = localStorage.getItem('last_modified_timestamp');
        const timestampSpan = document.getElementById('lastModifiedTimestamp');
        if (timestampSpan && savedTimestamp) {
            timestampSpan.textContent = savedTimestamp;
        }
    }
    let autoDownloadIntervalId = null;
    const autoDownloadStartTime = 7 * 60 + 30;
    const autoDownloadEndTime = 18 * 60;
    const autoDownloadInterval = 60 * 60 * 1000;

// === REPLACE showCustomModal con questa versione ===
function showCustomModal(title, message, buttons, inputConfig = null, selectConfig = null) {
        return new Promise(resolve => {
            // Imposta titolo e messaggio (consente anche HTML semplice nel messaggio)
            modalTitle.textContent = title || '';
            modalMessage.innerHTML = message || '';

            // Pulisce eventuali contenuti dinamici precedenti
            modalButtons.innerHTML = '';
            const existingInput = document.getElementById('modalInput');
            if (existingInput) existingInput.remove();
            const existingSelect = document.getElementById('modalSelect');
            if (existingSelect) existingSelect.remove();

            // Funzione per pulire e risolvere la promise
            const cleanup = (result) => {
                customModal.classList.remove('visible');
                document.body.classList.remove('modal-open');
                document.removeEventListener('keydown', onKeyDown);
                // rimuove input/select creati
                const i2 = document.getElementById('modalInput'); if (i2) i2.remove();
                const s2 = document.getElementById('modalSelect'); if (s2) s2.remove();
                resolve(result);
            };
            // Listener per il tasto ESC
            const onKeyDown = (e) => {
                if (e.key === 'Escape') {
                    cleanup(null);
                }
            };

            // Inserisce campo input opzionale
            if (inputConfig) {
                const inputHtml = `<input type="text" id="modalInput" placeholder="${inputConfig.placeholder || ''}" style="width: 90%; padding: 8px; margin: 10px 0 15px; border: 1px solid #ccc; border-radius: 5px;">`;
                modalMessage.insertAdjacentHTML('afterend', inputHtml);
            }
            // Inserisce select opzionale
            if (selectConfig && Array.isArray(selectConfig.options) && selectConfig.options.length > 0) {
                const selectHtml = `
        <select id="modalSelect" style="width: 90%; padding: 8px; margin: 10px 0 15px; border: 1px solid #ccc; border-radius: 5px;">
          ${selectConfig.options.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
        </select>`;
                modalMessage.insertAdjacentHTML('afterend', selectHtml);
            }

            // Crea bottoni basandosi sulla configurazione
            buttons.forEach(btnConfig => {
                const button = document.createElement('button');
                button.classList.add('modal-button', btnConfig.class || 'confirm');
                button.textContent = btnConfig.text || 'OK';
                button.onclick = () => {
                    let result = btnConfig.value;
                    if (inputConfig) {
                        result = (btnConfig.value === true) ? document.getElementById('modalInput').value : null;
                    } else if (selectConfig) {
                        result = (btnConfig.value === true) ? document.getElementById('modalSelect').value : btnConfig.value;
                    }
                    cleanup(result);
                };
                modalButtons.appendChild(button);
            });

            // Mostra il modale e prepara gli eventi
            customModal.classList.add('visible');
            document.body.classList.add('modal-open');
            document.addEventListener('keydown', onKeyDown);
            // Previene lo scroll della pagina sottostante
            customModal.addEventListener('wheel', (e) => { e.stopPropagation(); }, { passive: true });

            // Imposta il focus sull'input (se presente) o sul primo bottone
            const inp = document.getElementById('modalInput');
            if (inp) inp.focus();
            else {
                const firstBtn = modalButtons.querySelector('button');
                if (firstBtn) firstBtn.focus();
            }
        });
    }

    function showAlert(message, title = 'Attenzione') {
        return showCustomModal(title, message, [{ text: 'OK', class: 'alert', value: true }]);
    }

    function showConfirm(message, title = 'Conferma') {
        return showCustomModal(title, message, [
            { text: 'Sì', class: 'confirm', value: true },
            { text: 'No', class: 'cancel', value: false }
        ]);
    }

    async function showPromptModal(title, message, placeholder = '') {
        return await showCustomModal(title, message, [
            { text: 'Salva', class: 'confirm', value: true },
            { text: 'Annulla', class: 'cancel', value: null }
        ], { placeholder: placeholder });
    }

    async function showSelectionModal(title, message, options) {
        return await showCustomModal(title, message, [
            { text: 'Carica', class: 'confirm', value: true },
            { text: 'Elimina Selezionato', class: 'delete', value: 'delete' },
            { text: 'Annulla', class: 'cancel', value: null }
        ], null, { options: options });
    }

    function getWeekNumber(d) {
        d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
        d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
        const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
        const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
        return weekNo;
    }

    const today = new Date();
    currentDateSpan.textContent = today.toLocaleDateString('it-IT');
    currentWeekSpan.textContent = getWeekNumber(today);

    /*
     * Elenco dei macchinari disponibili per la selezione nelle tabelle di
     * produzione.  Oltre ai turboemulsori di varie dimensioni, vengono
     * aggiunti "Confezionamento" e "Spazio Libero" in modo da coprire
     * tutte le opzioni richieste dall'utente.  Questo array viene
     * utilizzato per generare i datalist sia nella tabella dettagli di
     * produzione sia nel programma giornaliero.  L'utente può sempre
     * inserire testo libero: il datalist serve solo come suggerimento.
     */
    const macchinariOptions = [
        "Turboemulsore 5L",
        "Turboemulsore 10L",
        "Turboemulsore 50L",
        "Turboemulsore 150L",
        "Turboemulsore 500L",
        "Miscelatore",
        "Camera Bianca",
        "Turbina Piccola",
        "Fusore",
        "Confezionamento",
        "Spazio Libero"
    ];

    const siNoOptions = ['si', 'no', ''];

    function isMedicalDeviceCode(code) {
        const codeStr = String(code || '').trim().toUpperCase();
        const specificMedicalDeviceCodes = ['7545', '40125V', '7316', '7317'];
        const is4xxxxCode = codeStr.startsWith('4') && codeStr.length >= 4 && /^\d+$/.test(codeStr.replace('V', ''));
        const isSpecificCode = specificMedicalDeviceCodes.includes(codeStr);
        // Includi anche i codici presenti nel DeviceRef importato come dispositivi medici
        let isException = false;
        try {
            const deviceRefs = JSON.parse(localStorage.getItem('deviceRefData') || '[]');
            if (Array.isArray(deviceRefs)) {
                isException = deviceRefs.some(ref => String(ref.codice || '').trim().toUpperCase() === codeStr);
            }
        } catch (e) {
            // se non disponibile, ignora l'eccezione
            isException = false;
        }
        return is4xxxxCode || isSpecificCode || isException;
    }

    function checkProductionNecessity(row) {
        const qtyRequestedInput = row.querySelector('.qty-requested-input');
        const stockInput = row.querySelector('.stock-input');
        const productionFlag = row.querySelector('.production-flag');

        const quantitaRichiesta = parseFloat(qtyRequestedInput.value);
        const giacenzaMagazzino = parseFloat(stockInput.value);

        if (!isNaN(quantitaRichiesta) && !isNaN(giacenzaMagazzino) && giacenzaMagazzino >= quantitaRichiesta) {
            productionFlag.textContent = '✅';
            productionFlag.classList.add('visible');
        } else {
            productionFlag.textContent = '✅';
            productionFlag.classList.remove('visible');
        }
    }

    function setupSiNoSelect(selectElement) {
        const updateStyle = () => {
            selectElement.classList.remove('si', 'no');
            if (selectElement.value === 'si') {
                selectElement.classList.add('si');
            } else if (selectElement.value === 'no') {
                selectElement.classList.add('no');
            }
        };
        selectElement.addEventListener('change', updateStyle);
        updateStyle();
    }

    function validateInput(inputElement, feedbackSpan, condition, message, isError = false) {
        if (condition) {
            if (feedbackSpan) {
                feedbackSpan.textContent = '';
                feedbackSpan.title = '';
            }
            inputElement.classList.remove('invalid-input-highlight');
        } else {
            if (feedbackSpan) {
                feedbackSpan.textContent = '';
                feedbackSpan.title = '';
            }
            inputElement.classList.remove('invalid-input-highlight');
        }
    }

    function validateRow(row) {
        checkProductionNecessity(row);

        const qtyRequestedInput = row.querySelector('.qty-requested-input');
        const stockInput = row.querySelector('.stock-input');
        const qtyToProduceInput = row.querySelector('.qty-to-produce-input');
        const packagingPiecesInput = row.querySelector('.packaging-pieces-input');
        const packagingKgPerPieceInput = row.querySelector('.packaging-kg-per-piece-input');
        const codeInput = row.querySelector('.code-input');
        const productionDaysInput = row.querySelector('.production-days-input');
        const machineInput = row.querySelector('.machine-input');
        const packagingUnitSelect = row.querySelector('.col-confez-kg-pezzo .unit-select');

        const qtyRequested = parseFloat(qtyRequestedInput.value);
        const stock = parseFloat(stockInput.value);
        const qtyToProduce = parseFloat(qtyToProduceInput.value);
        const packagingPieces = parseFloat(packagingPiecesInput.value);
        const packagingKgPerPiece = parseFloat(packagingKgPerPieceInput.value);
        const codeValue = codeInput.value.trim();
        const productionDays = parseFloat(productionDaysInput.value);

        const needsProduction = isNaN(qtyRequested) || isNaN(stock) || stock < qtyRequested;

        machineInput.value = assignMachine(codeValue, qtyToProduce, qtyRequested, stock);

        if (isMedicalDeviceCode(codeValue)) {
            packagingUnitSelect.value = 'mL';
        }

        const qtyToProduceFeedback = row.querySelector('.validation-feedback[data-for="qty-to-produce-input"]');
        validateInput(qtyToProduceInput, qtyToProduceFeedback, !needsProduction || (!isNaN(qtyToProduce) && qtyToProduce > 0), 'Quantità da Produrre non valida o mancante se la produzione è necessaria.', true);

        const packagingPiecesFeedback = row.querySelector('.validation-feedback[data-for="packaging-pieces-input"]');
        if (packagingPiecesInput.value.trim() !== '') {
            validateInput(packagingPiecesInput, packagingPiecesFeedback, !isNaN(packagingPieces) && packagingPieces >= 0, 'Numero Pezzi deve essere un valore numerico non negativo.', true);
        } else {
            validateInput(packagingPiecesInput, packagingPiecesFeedback, true, '');
        }

        const packagingKgPerPieceFeedback = row.querySelector('.validation-feedback[data-for="packaging-kg-per-piece-input"]');
        if (packagingKgPerPieceInput.value.trim() !== '') {
            validateInput(packagingKgPerPieceInput, packagingKgPerPieceFeedback, !isNaN(packagingKgPerPiece) && packagingKgPerPiece >= 0, 'Kg/Pezzo deve essere un valore numerico non negativo.', true);
        } else {
            validateInput(packagingKgPerPieceInput, packagingKgPerPieceFeedback, true, '');
        }

        const codeFeedback = row.querySelector('.validation-feedback[data-for="code-input"]');
        if (!/^\d{1,10}$/.test(codeValue) && !isMedicalDeviceCode(codeValue)) {
            validateInput(codeInput, codeFeedback, false, 'Codice deve essere un numero di massimo 10 cifre o un codice dispositivo medico specifico.', true);
        } else {
            validateInput(codeInput, codeFeedback, true, '');
        }

        const productionDaysFeedback = row.querySelector('.validation-feedback[data-for="production-days-input"]');
        if (productionDaysInput.value.trim() !== '') {
            validateInput(productionDaysInput, productionDaysFeedback, !isNaN(productionDays) && productionDays > 0, 'Giorni di Produzione deve essere un numero positivo.', true);
        } else {
            validateInput(productionDaysInput, productionDaysFeedback, true, '');
        }
    }

    function assignMachine(code, quantityToProduce, qtyRequested, stock) {
        const codeString = String(code || '').trim();
        const qty = parseFloat(quantityToProduce);
        const requested = parseFloat(qtyRequested);
        const currentStock = parseFloat(stock);

        if (!isNaN(requested) && !isNaN(currentStock) && currentStock >= requested) {
            return "";
        }

        if (isMedicalDeviceCode(codeString)) {
            return "Camera Bianca";
        }

        if (isNaN(qty)) {
            return "";
        } else if (qty >= 0 && qty <= 5) {
            return "Turboemulsore 5L";
        } else if (qty >= 6 && qty <= 10) {
            return "Turboemulsore 10L";
        } else if (qty >= 11 && qty <= 50) {
            return "Turboemulsore 50L";
        } else if (qty >= 51 && qty <= 150) {
            return "Turboemulsore 150L";
        } else if (qty > 150) {
            return "Turboemulsore 500L";
        }
        return "";
    }

    const parseNumericValue = (value) => {
        if (typeof value === 'number') {
            return value;
        }
        if (typeof value === 'string') {
            const match = value.match(/(\d+([.,]\d+)?)/);
            if (match) {
                const parsed = parseFloat(match[1].replace(',', '.'));
                if (!isNaN(parsed)) {
                    return parsed;
                }
            }
        }
        return '';
    };

    const parseDateValue = (value) => {
        if (value instanceof Date) {
            return value.toLocaleDateString('it-IT');
        }
        if (typeof value === 'string') {
            value = value.trim();
            let parts = value.split(/[\/\.]/);
            if (parts.length === 3) {
                const day = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10) - 1;
                let year = parseInt(parts[2], 10);

                if (year < 100) {
                    year = (year > 50) ? (1900 + year) : (2000 + year);
                }

                const date = new Date(year, month, day);
                if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
                    return date.toLocaleDateString('it-IT');
                }
            }
            parts = value.split('-');
            if (parts.length === 3) {
                const year = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10) - 1;
                const day = parseInt(parts[2], 10);
                const date = new Date(year, month, day);
                if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
                    return date.toLocaleDateString('it-IT');
                }
            }
        }
        return '';
    };

    const normalizeUnit = (unitStr) => {
        if (!unitStr) return '';
        const lowerUnit = unitStr.toLowerCase();
        if (lowerUnit === 'kg' || lowerUnit === 'kgs') return 'Kg';
        if (lowerUnit === 'pz' || lowerUnit === 'pezzi') return 'Pz';
        if (lowerUnit === 'ml' || lowerUnit === 'millilitri') return 'mL';
        if (lowerUnit === 'l' || lowerUnit === 'litri') return 'L';
        if (lowerUnit === 'g' || lowerUnit === 'grammi') return 'g';
        return unitStr;
    };

    const parsePackagingString = (packagingStr) => {
        let pezzi = '';
        let kgPerPezzo = '';
        let unit = '';
        let rawString = '';

        if (typeof packagingStr === 'string') {
            rawString = packagingStr.trim();
            const matchX = rawString.match(/^(\d+(?:[.,]\d+)?)[Xx](\d+(?:[.,]\d+)?)([a-zA-Z]+)?$/);
            const matchPiecesOnly = rawString.match(/^(\d+(?:[.,]\d+)?)\s*pz$/i);
            const matchSingleValue = rawString.match(/^(\d+(?:[.,]\d+)?)([a-zA-Z]+)?$/);

            if (matchX) {
                pezzi = parseNumericValue(matchX[1]);
                kgPerPezzo = parseNumericValue(matchX[2]);
                unit = normalizeUnit(matchX[3] || 'Kg');
            } else if (matchPiecesOnly) {
                pezzi = parseNumericValue(matchPiecesOnly[1]);
                unit = 'Pz';
            } else if (matchSingleValue) {
                const val = parseNumericValue(matchSingleValue[1]);
                const unitMatch = normalizeUnit(matchSingleValue[2] || '');
                if (unitMatch) {
                    pezzi = val;
                    unit = unitMatch;
                } else {
                    pezzi = val;
                }
            }
        } else if (typeof packagingStr === 'number') {
            pezzi = packagingStr;
        }

        return {
            pezzi: pezzi,
            kgPerPezzo: kgPerPezzo,
            unit: unit || '',
        };
    };

    function updateMaterialStatusFromNotes(row) {
        const notesInput = row.querySelector('.notes-input');
        const materiePrimeSelect = row.querySelector('.materie-prime-select');
        const materialeConfezSelect = row.querySelector('.materiale-confez-select');

        const notes = (notesInput.value || '').toLowerCase();

        let materiePrimeStatus = 'si';
        let materialeConfezionamentoStatus = 'si';

        const packagingKeywords = [
            'no astuccio', 'no blister', 'no bugiardino', 'no box', 'no etichetta',
            'senza astuccio', 'senza blister', 'senza bugiardino', 'senza box', 'senza etichetta',
            'manca astuccio', 'manca blister', 'manca bugiardino', 'manca box', 'manca etichetta',
            'no scatola', 'senza scatola', 'manca scatola',
            'no triplo', 'senza triplo', 'manca triplo',
            'no siringa', 'senza siringa', 'manca siringa',
            'no fiala', 'senza fiala', 'manca fiala'
        ];
        const foundPackagingKeyword = packagingKeywords.some(keyword => notes.includes(keyword));

        const rawMaterialKeywords = [
            'no materia prima', 'no materie prime', 'no vitamina', 'no vitamine','no mp',
            'senza materia prima', 'senza materie prime', 'senza vitamina', 'senza vitamine',
            'manca materia prima', 'manca materie prime', 'manca vitamina', 'manca vitamine'
        ];
        const foundRawMaterialKeyword = rawMaterialKeywords.some(keyword => notes.includes(keyword));

        const genericNoPattern = /no\s+\(.+?\)/;
        const genericNoMatch = notes.match(genericNoPattern);
        let foundGenericNonPackagingNo = false;
        if (genericNoMatch) {
            const genericTerm = genericNoMatch[0].replace(/no\s+\(|\)/g, '').trim();
            const isPackagingRelated = packagingKeywords.some(keyword => genericTerm.includes(keyword.replace(/no\s+/,'')));
            if (!isPackagingRelated) {
                foundGenericNonPackagingNo = true;
            }
        }

        if (foundPackagingKeyword) {
            materialeConfezionamentoStatus = 'no';
        }

        if (foundRawMaterialKeyword || foundGenericNonPackagingNo) {
            materiePrimeStatus = 'no';
        }

        materiePrimeSelect.value = materiePrimeStatus;
        materialeConfezSelect.value = materialeConfezionamentoStatus;

        materiePrimeSelect.dispatchEvent(new Event('change'));
        materialeConfezSelect.dispatchEvent(new Event('change'));
    }

    function createRow(rowData = {}) {
        rowData.materiePrime = rowData.materiePrime !== undefined ? rowData.materiePrime : '';
        rowData.materialeConfezionamento = rowData.materialeConfezionamento !== undefined ? rowData.materialeConfezionamento : '';
        // Di default l'unità di misura è impostata a "Pz" (pezzi) invece di Kg per rendere coerente l'unità
        rowData.confezionamentoUnit = rowData.confezionamentoUnit || 'Pz';
        rowData.quantitaRichiestaUnit = rowData.quantitaRichiestaUnit || 'Pz';

        const parsedQtyRequested = parseFloat(rowData.quantitaRichiesta);
        const parsedGiacenza = parseFloat(rowData.giacenzaMagazzino);
        let calculatedQtyToProduce = rowData.quantitaDaProdurre;

        const needsProduction = isNaN(parsedQtyRequested) || isNaN(parsedGiacenza) || parsedGiacenza < parsedQtyRequested;

        if (needsProduction && (isNaN(parseFloat(rowData.quantitaDaProdurre)) || parseFloat(rowData.quantitaDaProdurre) <= 0)) {
            calculatedQtyToProduce = Math.max(0, parsedQtyRequested - parsedGiacenza);
        }
        if (calculatedQtyToProduce < 0) {
            calculatedQtyToProduce = 0;
        }
        calculatedQtyToProduce = isNaN(calculatedQtyToProduce) || calculatedQtyToProduce === null ? '' : calculatedQtyToProduce;

        rowData.macchinari = assignMachine(rowData.codice, calculatedQtyToProduce, parsedQtyRequested, parsedGiacenza);

        const codeString = String(rowData.codice || '').trim();
        const isMedicalDevice = isMedicalDeviceCode(codeString);

        const row = document.createElement('tr');
        if (isMedicalDevice) {
            row.classList.add('highlight-code-4');
        }

        const rawPackagingDetailString = rowData.rawConfezionamentoDettaglio ||
            (rowData.confezionamentoPezzi && rowData.confezionamentoKgPerPiece && rowData.confezionamentoUnit ?
                `${rowData.confezionamentoPezzi}X${rowData.confezionamentoKgPerPiece}${rowData.confezionamentoUnit}` : '');

        const packagingDetails = parsePackagingString(rawPackagingDetailString);

        rowData.confezionamentoPezzi = packagingDetails.pezzi;
        rowData.confezionamentoKgPerPiece = packagingDetails.kgPerPezzo;
        rowData.confezionamentoUnit = packagingDetails.unit;

        if (isMedicalDeviceCode(rowData.codice)) {
            rowData.confezionamentoUnit = 'mL';
        }

        row.innerHTML = `
            <td><input type="checkbox" class="row-selector"></td>
            <td class="col-codice">
                <div class="input-with-unit">
                    <input type="text" class="code-input" value="${rowData.codice || ''}" maxlength="10">
                    <span class="validation-feedback" data-for="code-input"></span>
                </div>
            </td>
            <td class="col-prodotto"><input type="text" class="product-input" value="${rowData.prodotto || ''}" maxlength="56"></td>
            <td class="col-cliente"><input type="text" class="client-input" value="${rowData.cliente || ''}"></td>
            <td class="col-qty-richiesta">
                <div class="input-with-unit">
                    <input type="number" class="qty-requested-input" value="${rowData.quantitaRichiesta || ''}" min="0" max="9999">
                    <select class="unit-select qty-requested-unit-select">
                        ${unitOptions.map(opt => `<option value="${opt}" ${rowData.quantitaRichiestaUnit === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                    </select>
                    <span class="production-flag">✅</span>
                </div>
            </td>
            <td class="col-giacenza">
                <div class="input-with-unit">
                    <input type="number" class="stock-input" value="${rowData.giacenzaMagazzino || ''}" min="0" max="9999">
                    <span class="unit-label">Kg</span>
                </div>
            </td>
            <td class="col-qty-da-produrre">
                <div class="input-with-unit">
                    <input type="number" class="qty-to-produce-input" value="${calculatedQtyToProduce}" min="0" max="9999">
                    <span class="unit-label">Kg</span>
                    <span class="validation-feedback" data-for="qty-to-produce-input"></span>
                </div>
            </td>
            <td class="col-materie-prime">
                <select class="si-no-select materie-prime-select">
                    <option value="" ${rowData.materiePrime === '' ? 'selected' : ''}>Seleziona</option>
                    <option value="no" ${rowData.materiePrime === 'no' ? 'selected' : ''}>No</option>
                    <option value="si" ${rowData.materiePrime === 'si' ? 'selected' : ''}>Sì</option>
                </select>
            </td>
            <td class="col-macchinari">
                <input type="text" class="machine-input" list="macchinariOptionsList" value="${rowData.macchinari || ''}">
                <datalist id="macchinariOptionsList">
                    <!-- Opzione vuota per permettere inserimento libero -->
                    <option value=""></option>
                    ${macchinariOptions.map(opt => `<option value="${opt}"></option>`).join('')}
                </datalist>
            </td>
            <td class="col-operatore"><input type="text" class="operator-input" value="${rowData.operatore || ''}"></td>
            <td class="col-confez-pezzi">
                <div class="input-with-unit">
                    <input type="number" class="packaging-pieces-input" value="${rowData.confezionamentoPezzi || ''}" min="0" max="9999">
                    <span class="validation-feedback" data-for="packaging-pieces-input"></span>
                </div>
            </td>
            <td class="col-confez-kg-pezzo">
                <div class="input-with-unit">
                    <input type="number" class="packaging-kg-per-piece-input" value="${rowData.confezionamentoKgPerPiece || ''}" min="0" max="9999">
                    <select class="unit-select">
                        ${unitOptions.map(opt => `<option value="${opt}" ${rowData.confezionamentoUnit === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                    </select>
                    <span class="validation-feedback" data-for="packaging-kg-per-piece-input"></span>
                </div>
            </td>
            <td class="col-prod-data"><input type="text" class="datepicker production-date-input" placeholder="Seleziona data" value="${rowData.produzioneData || ''}"></td>
            <td class="col-giorni-produzione">
                <div class="input-with-unit">
                    <input type="number" class="production-days-input" value="${rowData.giorniDiProduzione || ''}" min="1" max="30">
                    <span class="validation-feedback" data-for="production-days-input"></span>
                </div>
            </td>
            <td class="col-data-confez"><input type="text" class="datepicker packaging-date-input" placeholder="Seleziona data" value="${rowData.dataConfezionamento || ''}"></td>
            <td class="col-cod-confez"><input type="text" class="packaging-code-input" value="${rowData.codiceConfezionamento || ''}"></td>
            <td class="col-lotto-sc"><input type="text" class="lotto-sc-input" value="${rowData.lottoSC || ''}"></td>
            <td class="col-materiale-confez">
                <select class="si-no-select materiale-confez-select">
                    <option value="" ${rowData.materialeConfezionamento === '' ? 'selected' : ''}>Seleziona</option>
                    <option value="no" ${rowData.materialeConfezionamento === 'no' ? 'selected' : ''}>No</option>
                    <option value="si" ${rowData.materialeConfezionamento === 'si' ? 'selected' : ''}>Sì</option>
                </select>
            </td>
            <td class="col-data-sped"><input type="text" class="datepicker shipping-date-input" placeholder="Seleziona data" value="${rowData.dataSpedizione || ''}"></td>
            <td class="col-note"><input type="text" class="notes-input" value="${rowData.note || ''}"></td>
        `;

        const datepickers = row.querySelectorAll('.datepicker');
        datepickers.forEach(input => {
            flatpickr(input, {
                dateFormat: "d/m/Y",
                locale: "it",
            });
        });

        const inputsToValidate = row.querySelectorAll('.code-input, .qty-requested-input, .stock-input, .qty-to-produce-input, .packaging-pieces-input, .packaging-kg-per-piece-input, .code-input, .production-days-input');
        inputsToValidate.forEach(input => {
            input.addEventListener('input', () => validateRow(row));
            input.addEventListener('change', () => validateRow(row));
        });

        const unitSelects = row.querySelectorAll('.unit-select');
        unitSelects.forEach(select => {
            select.addEventListener('change', () => validateRow(row));
        });

        const materiePrimeSelect = row.querySelector('.materie-prime-select');
        const materialeConfezSelect = row.querySelector('.materiale-confez-select');
        setupSiNoSelect(materiePrimeSelect);
        setupSiNoSelect(materialeConfezSelect);

        const notesInput = row.querySelector('.notes-input');
        notesInput.addEventListener('input', () => updateMaterialStatusFromNotes(row));
        notesInput.addEventListener('change', () => updateMaterialStatusFromNotes(row));


        const codeInput = row.querySelector('.code-input');
        const packagingKgPerPieceInput = row.querySelector('.packaging-kg-per-piece-input');
        const packagingCodeInput = row.querySelector('.packaging-code-input');

        const updatePackagingCode = () => {
            const code = codeInput.value.trim();
            const kgPerPiece = packagingKgPerPieceInput.value.trim();
            const isMedical = isMedicalDeviceCode(code);

            // Se il codice è presente nella mappa dei codici sterili, usa
            // direttamente il valore predefinito dalla mappa.  In caso
            // contrario, applica la logica originale (codice* per dispositivi
            // medici, codice-kg per pezzo se specificato oppure semplicemente
            // il codice).
            let generatedCode = '';
            if (code) {
                if (sterilePackagingMap.hasOwnProperty(code)) {
                    generatedCode = sterilePackagingMap[code];
                } else if (isMedical) {
                    generatedCode = `${code}*`;
                } else if (kgPerPiece) {
                    generatedCode = `${code}-${kgPerPiece}`;
                } else {
                    generatedCode = code;
                }
            }
            packagingCodeInput.value = generatedCode;
        };

        updatePackagingCode();
        codeInput.addEventListener('input', updatePackagingCode);
        packagingKgPerPieceInput.addEventListener('input', updatePackagingCode);

        const codiceCell = row.querySelector('.col-codice');
        codiceCell.dataset.productName = rowData.prodotto || '';
        codiceCell.addEventListener('mouseover', (e) => {
            const productName = e.currentTarget.dataset.productName;
            if (productName) {
                showGenericTooltip(`<strong>Prodotto:</strong> ${productName}`, e);
            }
        });
        codiceCell.addEventListener('mouseout', hideGenericTooltip);

        if (rowData.quantitaRichiesta !== undefined && rowData.giacenzaMagazzino !== undefined) {
            validateRow(row);
        }

        updateMaterialStatusFromNotes(row);

        return row;
    }

    addRowBtn.addEventListener('click', async () => {
        const newRow = createRow();
        productionTableBody.appendChild(newRow);
        applyFilter();
        validateRow(newRow);
        updateScrollButtons();
        runFullCheck();
        updateWarehouseGanttChart();
        addInlineScrollbarsToWarehouseGantt(); 
        updateAnalisiTable();
        addLogEntry(`Aggiunta riga manuale.`);
        autoSaveAllData();
    });

let isSaving = false; // Flag per prevenire salvataggi sovrapposti

  // --- Helpers globali necessari ---

function autoSaveAllData() {
  try {
    if (typeof getAllTableData === 'function') {
      localStorage.setItem('production_data_autosave', JSON.stringify(getAllTableData()));
    }
    if (typeof getAllSalesOrderData === 'function') {
      localStorage.setItem('sales_order_data_autosave', JSON.stringify(getAllSalesOrderData()));
    }
    if (typeof getAllShippingData === 'function') {
      localStorage.setItem('shipping_schedule_data_autosave', JSON.stringify(getAllShippingData()));
    }
    if (typeof getAllArrivalData === 'function') {
      localStorage.setItem('arrival_schedule_data_autosave', JSON.stringify(getAllArrivalData()));
    }
    if (typeof getAllOverdueArrivalData === 'function') {
      localStorage.setItem('overdue_arrival_data_autosave', JSON.stringify(getAllOverdueArrivalData()));
    }
    // Salva anche i dati della tabella "Merce in Scadenza" per l'autosave
    if (typeof getAllExpiringGoodsData === 'function') {
      try {
        localStorage.setItem('expiring_goods_data_autosave', JSON.stringify(getAllExpiringGoodsData()));
      } catch (e) {
        console.warn('autoSaveAllData: impossibile salvare expiring goods:', e);
      }
    }
    // Salva anche i dati della quarantena affinché vengano mantenuti in caso di refresh
    if (typeof getAllQuarantineData === 'function') {
      try {
        localStorage.setItem('quarantine_data_autosave', JSON.stringify(getAllQuarantineData()));
      } catch (e) {
        console.warn('autoSaveAllData: impossibile salvare la quarantena:', e);
      }
    }
  } catch (e) {
    console.warn('autoSaveAllData error:', e);
  }
}

// Alcuni punti del codice chiamano questo: metto uno shim sicuro.
function filterAndRenderLogbook() {
  try {
    if (typeof renderLogbook === 'function') renderLogbook();
  } catch (e) {
    console.warn('filterAndRenderLogbook shim:', e);
  }
}
 

 

async function saveDataToServer() {
    if (isSaving) {
        console.log("Salvataggio già in corso, attendo...");
        return;
    }
    isSaving = true;
    console.log("Avvio salvataggio dati sul server...");

    const dataToSave = {
        production_data: getAllTableData(),
        sales_order_data: getAllSalesOrderData(),
        shipping_data: getAllShippingData(),
        arrival_data: getAllArrivalData(),
        medical_device_data: getAllMedicalDeviceData(),
        // Dati di produzione medical device importati da file esterni
        medical_production_data: getMedicalProductionData(),
        overdue_arrival_data: getAllOverdueArrivalData(), // <-- DATO AGGIUNTO
        // Includiamo anche i dati OPI dal localStorage affinché vengano salvati sul server
        opi_monitor_data: getOpiMonitorData(),
        // Includiamo i dati DeviceRef, contenenti le informazioni aggiuntive per
        // dispositivi e medicali (aghi, siringhe, volumi, pesi).
        device_ref_data: getDeviceRefData(),
        // Dati della quarantena (merce evasa da arrivi)
        quarantine_data: typeof getAllQuarantineData === 'function' ? getAllQuarantineData() : [],
        // Dati della tabella Merce in Scadenza (Inventario)
        expiring_goods_data: typeof getAllExpiringGoodsData === 'function' ? getAllExpiringGoodsData() : []
    };

    try {
        const response = await fetch(apiEndpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(dataToSave)
        });
        if (response.ok) {
            console.log('Dati salvati con successo sul server.');
            updateLastModifiedTimestamp();
        } else {
            const errorData = await response.json();
            console.error('Errore nel salvataggio sul server:', errorData.message);
            showAlert(`Errore nel salvataggio dei dati: ${errorData.message}`);
        }
    } catch (error) {
        console.error('Errore di rete durante il salvataggio:', error);
        const loginOverlayEl = document.getElementById('loginOverlay');
        const overlayVisible = loginOverlayEl && loginOverlayEl.style.display !== 'none';
        const isLoggedIn = (typeof currentUserLevel !== 'undefined' && currentUserLevel > 0);
              // Se non è disponibile la connessione al server, registra l'errore ma
              // evita di infastidire l'utente con un pop-up persistente.  I dati
              // rimarranno salvati localmente e potranno essere sincronizzati
              // quando la connessione sarà ripristinata.
              if (isLoggedIn && !overlayVisible) {
                  console.warn('Errore di connessione con il server. Dati salvati solo in locale.');
                  // facoltativamente aggiungi una voce al logbook se disponibile
                  if (typeof addLogEntry === 'function') {
                    addLogEntry('Errore di connessione con il server: dati salvati solo in locale.');
                  }
              } else {
                  // Durante il login o se l'utente non è loggato, silenzia l'errore
                  console.warn('Salvataggio saltato (non loggato/overlay visibile):', error);
              }
    } finally {
        isSaving = false;
    }
}


async function loadDataFromServer() {
    try {
        const response = await fetch(apiEndpoint, { method: 'GET', cache: 'no-cache' });
        if (response.ok) {
            const text = await response.text();
            if (text) { 
                const data = JSON.parse(text);
                
                if (data.production_data) {
                    productionTableBody.innerHTML = '';
                    data.production_data.forEach(rowData => productionTableBody.appendChild(createRow(rowData)));
                }
                if (data.sales_order_data) {
                    salesOrderTableBody.innerHTML = '';
                    data.sales_order_data.forEach(rowData => salesOrderTableBody.appendChild(createSalesOrderRow(rowData)));
                }
                if (data.shipping_data) {
                    shippingScheduleTableBody.innerHTML = '';
                    data.shipping_data.forEach(rowData => shippingScheduleTableBody.appendChild(createShippingScheduleRow(rowData)));
                }
                if (data.arrival_data) {
                    arrivalScheduleTableBody.innerHTML = '';
                    data.arrival_data.forEach(rowData => arrivalScheduleTableBody.appendChild(createArrivalScheduleRow(rowData)));
                }
                      // Aggiorna la tabella produzione dispositivi medici solo se il server
                      // restituisce effettivamente dei dati non vuoti.  In assenza di
                      // dati (undefined o array vuota) si mantiene il contenuto
                      // locale, evitando di sovrascrivere un'importazione manuale.
                      if (Array.isArray(data.medical_device_data) && data.medical_device_data.length > 0 && medicalDeviceTableBody) {
                          medicalDeviceTableBody.innerHTML = '';
                          data.medical_device_data.forEach(rowData => medicalDeviceTableBody.appendChild(createMedicalDeviceRow(rowData)));
                      }
                // ===== LOGICA AGGIUNTA per caricare i dati non arrivati =====
                // Carica la tabella "Merce non Arrivata" solo se il server restituisce
                // dati non vuoti; in caso contrario mantiene i dati locali (se presenti)
                if (data.overdue_arrival_data && Array.isArray(data.overdue_arrival_data) && data.overdue_arrival_data.length > 0 && overdueArrivalsTableBody) {
                    populateOverdueTable(data.overdue_arrival_data);
                }
                // ===========================================================

                // ===== Caricamento dei dati OPI (Ordini di Produzione Interni) =====
                if (data.opi_monitor_data) {
                    // Aggiorna lo storage locale e popola la tabella OPI
                    localStorage.setItem('opi_monitor_data', JSON.stringify(data.opi_monitor_data));
                    populateOpiTable(data.opi_monitor_data);
                }
                // ===============================================================

                // ===== Caricamento dei dati DeviceRef (dispositivi/medicali) =====
                if (data.device_ref_data) {
                    // Aggiorna lo storage locale con i riferimenti dispositivi
                    try {
                        localStorage.setItem('deviceRefData', JSON.stringify(data.device_ref_data));
                    } catch (e) {
                        console.warn('Impossibile salvare i dati DeviceRef in localStorage:', e);
                    }
                }
                // ===== Caricamento dei dati di produzione medicale =====
                // Carica i dati solo se presenti e non vuoti.  In assenza di dati dal
                // server, mantiene le informazioni locali per preservare l'ultimo
                // import.  Questo impedisce che un import successivo o un reload
                // sovrascriva involontariamente i dati salvati localmente.
                if (data.medical_production_data && Array.isArray(data.medical_production_data) && data.medical_production_data.length > 0) {
                    try {
                        localStorage.setItem('medicalProductionData', JSON.stringify(data.medical_production_data));
                    } catch (e) {
                        console.warn('Impossibile salvare medicalProductionData in localStorage:', e);
                    }
                    populateMedicalDeviceProductionTable(data.medical_production_data);
                }
                // ===============================================================

                // ===== Caricamento dei dati di quarantena (merce evasa) =====
                // Aggiorna i dati di quarantena solo se la risposta del server
                // contiene un array non vuoto.  In caso contrario mantiene le
                // informazioni locali per evitare la perdita dei dati salvati.
                if (data.quarantine_data && Array.isArray(data.quarantine_data) && data.quarantine_data.length > 0) {
                    try {
                        // Memorizza i dati anche nel localStorage per un recupero rapido
                        localStorage.setItem('quarantine_data_autosave', JSON.stringify(data.quarantine_data));
                    } catch (e) {
                        console.warn('Impossibile salvare i dati della quarantena in localStorage:', e);
                    }
                    populateQuarantineTable(data.quarantine_data);
                }
                // ===============================================================

                console.log('Dati ricaricati dal server.');
                updateAllUIComponents();
            }
        }
    } catch (error) {
        console.error('Errore di rete durante il caricamento:', error);
    }
}

async function saveOpiDataToLocalAndServer(opiData) {
    localStorage.setItem('opi_monitor_data', JSON.stringify(opiData));
    // Salva anche i dati OPI sul server includendoli nei dati generali
    try {
        await saveDataToServer();
        // Aggiorna la data/ora dell'ultimo import OPI usando il formato leggibile.
        if (typeof formatDateTimeForDisplay === 'function') {
            const nowStr = formatDateTimeForDisplay(new Date());
            // Salva l'ultima importazione OPI su una chiave uniforme senza underscore
            localStorage.setItem('lastImportOPI', nowStr);
        } else {
            // Fallback se la funzione non esiste: salva l'Epoch time in stringa
            localStorage.setItem('lastImportOPI', Date.now().toString());
        }
        // Aggiorna le etichette di importazione e la sezione riassuntiva, se disponibile
        if (typeof updateImportTimestamps === 'function') {
            updateImportTimestamps();
        }
    } catch (e) {
        console.error('Errore nel salvataggio dei dati OPI sul server:', e);
    }
}
function loadOpiDataFromLocalAndServer() {
    const local = localStorage.getItem('opi_monitor_data');
    if (local) {
        try {
            const opiData = JSON.parse(local);
            populateOpiTable(opiData);
        } catch (e) {}
    }
}

    
    // Funzione helper per aggiornare tutti i componenti visivi dopo un caricamento
    function updateAllUIComponents() {
        updateGanttChart();
        updateWarehouseGanttChart();
        updateDailyProductionTable();
        updateAnalisiTable();
        updateMedicalDeviceProductionTable(); // <-- RIGA AGGIUNTA
        runFullCheck();

        // Dopo l'aggiornamento di tutte le componenti, controlla se esistono
        // spedizioni ADR e notifica l'utente se necessario.  Questo
        // richiamo assicura che l'avviso venga mostrato quando le spedizioni
        // vengono caricate dal server o dopo un'importazione.
        if (typeof checkAndNotifyADR === 'function') {
            try {
                checkAndNotifyADR();
            } catch (e) {
                console.warn('Errore nel controllo ADR:', e);
            }
        }

        // Dopo l'aggiornamento, controlla se esistono notifiche di magazzino
        // per la merce evasa.  Verrà mostrato un pop-up solo agli utenti CQ.
        if (typeof checkAndNotifyWarehouse === 'function') {
            try {
                checkAndNotifyWarehouse();
            } catch (e) {
                console.warn('Errore nel controllo magazzino:', e);
            }
        }

        // Dopo l'aggiornamento, controlla se esistono notifiche di spedizione
        // per gli ordini marcati come spediti.  Verrà mostrato un pop-up
        // solo agli utenti abilitati alle notifiche di spedizione.
        if (typeof checkAndNotifyShipping === 'function') {
            try {
                checkAndNotifyShipping();
            } catch (e) {
                console.warn('Errore nel controllo spedizioni:', e);
            }
        }
    }


// ===================================================================
// ==> NUOVE FUNZIONI PER GESTIRE I COMMENTI QA <==
// ===================================================================

/**
 * Gestisce il click sull'icona del lucchetto nei commenti QA.
 */
async function handleQACommentClick(event) {
    event.stopPropagation();
    const rowId = event.target.dataset.rowId;
    const targetRow = document.querySelector(`tr[data-row-id="${rowId}"]`);
    if (!targetRow) {
        showAlert("Errore: Riga di riferimento non trovata.");
        return;
    }

    // Modale per la password
    const passwordModal = document.createElement('div');
    passwordModal.className = 'modal-overlay visible';
    passwordModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Accesso Commenti QA</h3>
            <p>Inserisci la password per modificare i commenti.</p>
            <input type="password" id="qaPasswordInput" placeholder="Password...">
            <div class="modal-buttons">
                <button id="qaConfirmBtn" class="modal-button confirm">Conferma</button>
                <button id="qaCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(passwordModal);

    const qaConfirmBtn = document.getElementById('qaConfirmBtn');
    const qaCancelBtn = document.getElementById('qaCancelBtn');
    const qaPasswordInput = document.getElementById('qaPasswordInput');

    qaConfirmBtn.onclick = () => {
        if (qaPasswordInput.value === 'qa123') {
            passwordModal.remove();
            openQACommentsEditor(targetRow);
        } else {
            showAlert("Password non corretta.");
            qaPasswordInput.value = '';
        }
    };

    qaCancelBtn.onclick = () => passwordModal.remove();
    qaPasswordInput.focus();
}

    /**
     * Controlla se nelle spedizioni programmate sono presenti articoli ADR e,
     * in caso affermativo, visualizza un avviso all'utente.  L'avviso
     * indica il primo giorno in cui è prevista una spedizione ADR e il
     * codice articolo corrispondente.  L'utente può scegliere di
     * posporre l'avviso (verrà riproposto al prossimo caricamento) o di
     * confermarne la lettura (non verrà più mostrato fino a quando cambierà
     * la combinazione data/codice).  L'informazione persistente viene
     * salvata in localStorage con la chiave "adrNotificationAcknowledged".
     */
    function checkAndNotifyADR() {
        // Mostra un avviso per tutte le spedizioni ADR imminenti.  Il messaggio
        // viene creato dinamicamente includendo data, OV, codice, descrizione,
        // lotto e numero pezzi (ove disponibili).  La notifica viene
        // presentata solo agli utenti magazzino (livello 4) e commerciale
        // (livello 2) e supporta la gestione indipendente del consenso tramite
        // localStorage per ciascun ruolo.
        // Prima di proseguire verifica se l'utente è loggato e se il livello
        // utente corrente ha i permessi per ricevere le notifiche ADR (dato da
        // alertPermissions).  Se non siamo ancora loggati (currentUserLevel
        // non definito o 0), nascondi eventuali pop‑up residui e interrompi.
        if (!currentUserLevel || currentUserLevel <= 0) {
            const alertDiv = document.getElementById('adrNotification');
            if (alertDiv) alertDiv.style.display = 'none';
            return;
        }
        if (!alertPermissions[currentUserLevel] || !alertPermissions[currentUserLevel].ADR) {
            return;
        }
        if (typeof getAllShippingData !== 'function') return;
        const shippingData = getAllShippingData();
        if (!Array.isArray(shippingData) || shippingData.length === 0) return;
        // Determina il ruolo dell'utente per differenziare il tracciamento
        // dell'avviso nel localStorage.
        let roleSuffix = '';
        if (currentUserLevel === 4) {
            roleSuffix = 'magazzino';
        } else if (currentUserLevel === 2) {
            roleSuffix = 'commerciale';
        } else {
            // altri ruoli non ricevono avvisi ADR
            return;
        }
        // Filtra tutte le righe di spedizione con codice ADR
        const adrRows = shippingData.filter(row => {
            const code = (row.codiceArticolo || '').toString().trim().toUpperCase();
            return window.adrCodes && window.adrCodes.has(code);
        });
        if (adrRows.length === 0) {
            // Nessuna spedizione ADR imminente: assicurati che l'avviso sia nascosto
            const alertDiv = document.getElementById('adrNotification');
            if (alertDiv) alertDiv.style.display = 'none';
            return;
        }
        // Convertitore di date da formato italiano dd/mm/yyyy
        const parseItalianDate = (str) => {
            const parts = str.split('/');
            if (parts.length === 3) {
                const day = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10) - 1;
                const year = parseInt(parts[2], 10);
                return new Date(year, month, day);
            }
            return null;
        };
        const today = new Date();
        // Definisci un intervallo di 14 giorni a partire da oggi per includere
        // tutte le spedizioni in programma nella finestra visibile del gantt.
        const twoWeeksLater = new Date();
        twoWeeksLater.setDate(today.getDate() + 14);
        // Filtra le spedizioni ADR che cadranno nel prossimo intervallo (oggi
        // incluso).  Questo consente di segnalare tutte le spedizioni
        // imminenti e non solo la prima.
        const upcomingAdrRows = adrRows.filter(row => {
            const dt = parseItalianDate(row.dataConsegna);
            return dt && dt >= today && dt <= twoWeeksLater;
        });
        if (upcomingAdrRows.length === 0) {
            // Nessuna spedizione ADR da notificare: nascondi eventuale avviso rimasto
            const alertDiv = document.getElementById('adrNotification');
            if (alertDiv) alertDiv.style.display = 'none';
            return;
        }
        // Costruisci una chiave che rappresenta l'elenco corrente delle
        // spedizioni ADR.  In caso di modifica dell'elenco (nuove
        // spedizioni o cambiate le date/codici), l'avviso verrà riproposto.
        const adrKey = `${roleSuffix}_${upcomingAdrRows.map(r => `${r.ov || ''}-${r.codiceArticolo || ''}-${r.dataConsegna || ''}`).join('|')}`;
        const storedKey = localStorage.getItem('adrNotificationAcknowledged_' + roleSuffix);
        if (storedKey && storedKey === adrKey) {
            return; // l'utente ha già confermato per questo insieme di spedizioni
        }
        // Prepara il contenuto del messaggio.  Recupera eventuali informazioni
        // aggiuntive da opi_monitor_data (lotto, numero pezzi, um) per
        // arricchire l'avviso.
        let opiData = [];
        try {
            opiData = typeof getOpiMonitorData === 'function' ? getOpiMonitorData() : JSON.parse(localStorage.getItem('opi_monitor_data') || '[]');
        } catch (e) {
            opiData = [];
        }
        const listItemsHtml = upcomingAdrRows.map(row => {
            // Cerca lotto, quantità, UM e OP da opiData
            let lotto = '';
            let quantita = '';
            let um = '';
            let op = '';
            if (Array.isArray(opiData)) {
                const match = opiData.find(opi => {
                    return String(opi.ov || '').trim().toUpperCase() === String(row.ov || '').trim().toUpperCase() &&
                           String(opi.codice || '').trim().toUpperCase() === String(row.codiceArticolo || '').trim().toUpperCase();
                });
                if (match) {
                    lotto = match.lotto || '';
                    quantita = match.quantita || '';
                    um = match.um || '';
                    op = match.op || '';
                }
            }
            const descStr = row.descrizioneArticolo ? ` - ${row.descrizioneArticolo}` : '';
            // Mostra lotto sempre; se mancante visualizza N/D
            const lottoVal = lotto ? lotto : 'N/D';
            const lottoStr = ` - Lotto: ${lottoVal}`;
            let qtyStr = '';
            if (quantita) {
                qtyStr = ` - Numero pezzi: ${quantita}${um ? ' ' + um : ''}`;
            }
            const opStr = op ? ` - OP: ${op}` : '';
            return `<li><strong>${row.dataConsegna}</strong> - OV: <strong>${row.ov}</strong>${opStr} - Codice: <strong>${row.codiceArticolo}</strong>${descStr}${lottoStr}${qtyStr}</li>`;
        }).join('');
        // Recupera il container dell'avviso
        const alertDiv = document.getElementById('adrNotification');
        if (!alertDiv) return;
        const messageP = alertDiv.querySelector('p');
        const listEl = alertDiv.querySelector('ul');
        if (messageP) {
            messageP.innerHTML = `Attenzione: sono previste spedizioni ADR nei prossimi giorni:`;
        }
        if (listEl) {
            listEl.innerHTML = listItemsHtml;
        }
        alertDiv.style.display = 'block';
        // Gestione pulsanti
        const postponeBtn = document.getElementById('adrPostponeBtn');
        const okBtn = document.getElementById('adrAcknowledgeBtn');
        if (postponeBtn) {
            // Assicurati che i pulsanti siano visibili e cliccabili ogni volta
            postponeBtn.style.display = 'inline-flex';
            postponeBtn.disabled = false;
            postponeBtn.style.pointerEvents = 'auto';
            postponeBtn.onclick = () => {
                alertDiv.style.display = 'none';
            };
        }
        if (okBtn) {
            okBtn.style.display = 'inline-flex';
            okBtn.disabled = false;
            okBtn.style.pointerEvents = 'auto';
            okBtn.onclick = () => {
                localStorage.setItem('adrNotificationAcknowledged_' + roleSuffix, adrKey);
                alertDiv.style.display = 'none';
            };
        }

        // In ogni apertura dell'avviso ADR, collega il pulsante di chiusura
        // (la X nell'angolo in alto a destra) per consentire la
        // chiusura immediata del pop-up indipendentemente dai pulsanti
        // Posponi/OK.  Questo pulsante è sempre presente nel DOM.
        const adrCloseBtn = alertDiv.querySelector('.adr-close-btn');
        if (adrCloseBtn) {
            adrCloseBtn.style.display = 'inline';
            adrCloseBtn.style.pointerEvents = 'auto';
            adrCloseBtn.onclick = () => {
                alertDiv.style.display = 'none';
            };
        }
    }

    /**
     * Registra un cambiamento di stato da parte del CQ o del QA.  I cambiamenti
     * vengono memorizzati su server (via saveDataToServer) e localmente sotto
     * la chiave 'qualityStatusChanges'.  Per ciascuna combinazione
     * (type, code, lotto) viene mantenuta solo la versione più recente.
     * @param {string} type "CQ" o "QA"
     * @param {HTMLTableRowElement} targetRow la riga della tabella spedizioni su cui è stato modificato lo stato
     */
async function registerQualityStatusChange(type, targetRow) {
        try {
            let changes = [];
            try {
                changes = JSON.parse(localStorage.getItem('qualityStatusChanges') || '[]');
            } catch (e) {
                changes = [];
            }
            // Raccogli i dati principali dalla riga.
            const cells = targetRow.cells;
            const ov = cells[1] && cells[1].querySelector('input') ? cells[1].querySelector('input').value.trim() : '';
            const code = cells[2] && cells[2].querySelector('input') ? cells[2].querySelector('input').value.trim() : '';
            const descrizione = cells[3] && cells[3].querySelector('input') ? cells[3].querySelector('input').value.trim() : '';
            const newStatus = targetRow.dataset[type.toLowerCase() + 'Status'] || '';
            // Ricava lotto, quantità, unità di misura e OP dall'OPI monitor se possibile
            let lotto = '';
            let quantita = '';
            let um = '';
            let op = '';
            try {
                const opiData = typeof getOpiMonitorData === 'function' ? getOpiMonitorData() : JSON.parse(localStorage.getItem('opi_monitor_data') || '[]');
                if (Array.isArray(opiData)) {
                    const match = opiData.find(opi => {
                        return String(opi.ov || '').trim().toUpperCase() === String(ov || '').trim().toUpperCase() &&
                               String(opi.codice || '').trim().toUpperCase() === String(code || '').trim().toUpperCase();
                    });
                    if (match) {
                        lotto = match.lotto || '';
                        quantita = match.quantita || '';
                        um = match.um || '';
                        op = match.op || '';
                    }
                }
            } catch (e) {
                lotto = '';
            }
            const id = `${type}-${code}-${lotto}`;
            // Controlla se esiste già una modifica con lo stesso tipo, codice, lotto
            // e con lo stesso nuovo stato. In tal caso nessun aggiornamento è necessario.
            const existing = changes.find(item => item.type === type && item.code === code && item.lotto === lotto);
            if (existing && existing.newStatus === newStatus) {
                // Non salvare un record identico: evita notifiche duplicate
                return;
            }
            // Rimuovi tutte le versioni precedenti per questa combinazione type/code/lotto
            changes = changes.filter(item => !(item.type === type && item.code === code && item.lotto === lotto));
            // Inserisci la nuova versione aggiornata
            changes.push({
                id,
                type,
                code,
                descrizione,
                ov,
                op,
                lotto,
                newStatus,
                quantita,
                um,
                timestamp: new Date().toISOString()
            });
            localStorage.setItem('qualityStatusChanges', JSON.stringify(changes));
            // Salva il cambiamento anche sul server se possibile.
            // Usare await richiede che questa funzione sia dichiarata async.
            if (typeof saveDataToServer === 'function') {
                try {
                    // Salva le modifiche in modo asincrono; eventuali errori sono
                    // semplicemente registrati sul console senza impedire
                    // l'esecuzione.  Il server dovrà interpretare
                    // qualityStatusChanges dal localStorage.
                    await saveDataToServer();
                } catch (e) {
                    console.warn('Errore nel salvataggio su server dello stato CQ/QA:', e);
                }
            }
        } catch (err) {
            console.warn('Errore nel salvataggio del cambiamento CQ/QA:', err);
        }
    }

    // =====================================================================
    // ==> FUNZIONI PER LA MERCE IN QUARANTENA E LE NOTIFICHE MAGAZZINO <==
    // =====================================================================

    /**
     * Crea e aggiunge una nuova riga nella tabella "Merce in Quarantena".
     * I campi sono gli stessi della tabella "Merce non Arrivata"/"Arrivi".
     * Tutte le celle sono readonly per evitare modifiche accidentali.  La
     * selezione iniziale viene impostata tramite un checkbox nella prima colonna.
     * @param {Object} rowData Dati della riga da inserire
     */
    function createQuarantineRow(rowData = {}) {
        if (!quarantineTableBody) return null;
        const row = document.createElement('tr');
        // Funzione di escape per evitare problemi con le virgolette
        const esc = (str) => String(str || '').replace(/"/g, '&quot;');
        row.innerHTML = `
            <td><input type="checkbox" class="quarantine-row-selector"></td>
            <td><input type="text" value="${esc(rowData.ov || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.codiceArticolo || rowData.codice || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.descrizioneArticolo || rowData.descrizione || '')}" readonly style="text-align:left;"></td>
            <td><input type="text" value="${esc(rowData.layout || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.quantita || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.um || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.dataConsegna || rowData.data || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.dataConferma || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.ragioneSociale || rowData.cliente || '')}" readonly style="text-align:left;"></td>
            <td><input type="text" value="${esc(rowData.riferimentoCliente || '')}" readonly style="text-align:left;"></td>
            <td><input type="text" value="${esc(rowData.indirizzo || '')}" readonly style="text-align:left;"></td>
            <td><input type="text" value="${esc(rowData.cap || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.citta || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.provincia || '')}" readonly></td>
            <td><input type="text" value="${esc(rowData.telefono || '')}" readonly></td>
        `;
        quarantineTableBody.appendChild(row);
        return row;
    }

    /**
     * Estrae i dati da una riga della tabella quarantena restituendo un
     * oggetto identico a quelli usati per la tabella degli arrivi.
     * @param {HTMLTableRowElement} row
     */
    function getQuarantineRowData(row) {
        const cells = row.cells;
        return {
            ov: cells[1] && cells[1].querySelector('input') ? cells[1].querySelector('input').value : '',
            codiceArticolo: cells[2] && cells[2].querySelector('input') ? cells[2].querySelector('input').value : '',
            descrizioneArticolo: cells[3] && cells[3].querySelector('input') ? cells[3].querySelector('input').value : '',
            layout: cells[4] && cells[4].querySelector('input') ? cells[4].querySelector('input').value : '',
            quantita: cells[5] && cells[5].querySelector('input') ? cells[5].querySelector('input').value : '',
            um: cells[6] && cells[6].querySelector('input') ? cells[6].querySelector('input').value : '',
            dataConsegna: cells[7] && cells[7].querySelector('input') ? cells[7].querySelector('input').value : '',
            dataConferma: cells[8] && cells[8].querySelector('input') ? cells[8].querySelector('input').value : '',
            ragioneSociale: cells[9] && cells[9].querySelector('input') ? cells[9].querySelector('input').value : '',
            riferimentoCliente: cells[10] && cells[10].querySelector('input') ? cells[10].querySelector('input').value : '',
            indirizzo: cells[11] && cells[11].querySelector('input') ? cells[11].querySelector('input').value : '',
            cap: cells[12] && cells[12].querySelector('input') ? cells[12].querySelector('input').value : '',
            citta: cells[13] && cells[13].querySelector('input') ? cells[13].querySelector('input').value : '',
            provincia: cells[14] && cells[14].querySelector('input') ? cells[14].querySelector('input').value : '',
            telefono: cells[15] && cells[15].querySelector('input') ? cells[15].querySelector('input').value : ''
        };
    }

    /**
     * Restituisce tutti i dati presenti nella tabella quarantena come un array.
     */
    function getAllQuarantineData() {
        const rows = [];
        if (!quarantineTableBody) return rows;
        quarantineTableBody.querySelectorAll('tr').forEach(row => {
            rows.push(getQuarantineRowData(row));
        });
        return rows;
    }

    /**
     * Popola la tabella quarantena con un array di dati.  Tutto il contenuto
     * attuale viene rimosso prima di inserire le nuove righe.
     * @param {Array} data Array di oggetti da visualizzare
     */
    function populateQuarantineTable(data) {
        if (!quarantineTableBody) return;
        quarantineTableBody.innerHTML = '';
        if (Array.isArray(data)) {
            data.forEach(item => createQuarantineRow(item));
        }
        // Applica i filtri attivi dopo aver popolato la tabella.  Se la funzione
        // applyQuarantineFilters è definita (potrebbe non esserlo se il DOM non
        // è ancora caricato), richiamiamo per aggiornare la visibilità.
        if (typeof applyQuarantineFilters === 'function') {
            applyQuarantineFilters();
        }
    }

    /**
     * Gestisce il click sul pallino magazzino all'interno del Gantt degli arrivi.
     * Richiede la password ("mag345") la prima volta e permette di passare
     * dallo stato bianco (merce da evadere) a verde (merce evasa).  Quando
     * una riga viene marcata come evasa, viene rimossa dalla tabella degli
     * arrivi, aggiunta alla quarantena, salvata localmente e sul server, e
     * viene generata una notifica per gli utenti CQ.
     * @param {MouseEvent} event L'evento di click
     */
    async function handleMagStatusClick(event) {
        event.stopPropagation();
        const dot = event.target;
        const idxStr = dot.dataset.rowIndex;
        const index = parseInt(idxStr, 10);
        if (isNaN(index) || index < 0) return;
        const rows = arrivalScheduleTableBody ? arrivalScheduleTableBody.querySelectorAll('tr') : [];
        const row = rows[index];
        if (!row) return;
        // Richiede la password "mag345" ad ogni click sul pallino.  Non
        // utilizza piu' una variabile globale per memorizzare l'avvenuta
        // autenticazione, quindi l'operatore deve inserire la password per
        // ogni articolo evaso.  La verifica è case-sensitive.
        const password = window.prompt('Inserisci la password magazzino:');
        if (password !== 'mag345') {
            window.alert('Password errata.');
            return;
        }
        const currentStatus = row.dataset.magStatus || 'white';
        if (currentStatus === 'green') {
            // Se l'utente riclicca su una riga già evasa, passa allo stato bianco
            row.dataset.magStatus = 'white';
            dot.classList.remove('mag-status-green');
            dot.classList.add('mag-status-white');
            autoSaveAllData();
            updateWarehouseGanttChart();
            return;
        }
        // Aggiorna lo stato magazzino sulla riga
        row.dataset.magStatus = 'green';
        dot.classList.remove('mag-status-white');
        dot.classList.add('mag-status-green');
        // Copia i dati dalla riga di arrivo e inseriscili in quarantena
        const rowData = getArrivalScheduleRowData(row);
        createQuarantineRow(rowData);
        // Rimuove la riga dalla tabella degli arrivi
        arrivalScheduleTableBody.removeChild(row);
        // Aggiorna e salva i dati localmente e sul server
        autoSaveAllData();
        if (typeof saveDataToServer === 'function') {
            try {
                await saveDataToServer();
            } catch (e) {
                console.warn('Errore nel salvataggio dei dati dopo evasione:', e);
            }
        }
        // Registra la modifica di stato per mostrare la notifica al CQ
        registerWarehouseStatusChange(rowData);
        // Aggiorna il grafico di Gantt per riflettere la rimozione
        updateWarehouseGanttChart();
        // Mostra la notifica magazzino se necessario
        checkAndNotifyWarehouse();
    }

    /**
     * Registra un cambiamento di stato di magazzino (merce evasa) nel
     * localStorage.  Ogni entry contiene un identificatore univoco, i
     * principali dati della riga e un timestamp.  Le notifiche vengono
     * mostrate solo agli utenti con permesso CQ.
     * @param {Object} rowData Dati della riga evasa (ov, codice, descrizione, ecc.)
     */
    function registerWarehouseStatusChange(rowData) {
        try {
            let changes = [];
            try {
                changes = JSON.parse(localStorage.getItem('warehouseStatusChanges') || '[]');
            } catch (e) {
                changes = [];
            }
            const id = 'MAG-' + Date.now().toString(36) + '-' + Math.random().toString(36).substr(2, 5);
            // Raccogli le informazioni principali; alcune potrebbero non essere presenti
            const change = {
                id,
                code: rowData.codiceArticolo || rowData.codice || '',
                descrizione: rowData.descrizioneArticolo || rowData.descrizione || '',
                ov: rowData.ov || '',
                quantita: rowData.quantita || '',
                um: rowData.um || '',
                newStatus: 'Evasa',
                timestamp: new Date().toISOString()
            };
            // Inserisci la nuova notifica
            changes.push(change);
            localStorage.setItem('warehouseStatusChanges', JSON.stringify(changes));
        } catch (e) {
            console.warn('Impossibile registrare la modifica magazzino:', e);
        }
    }

    /**
     * Gestisce il click sulla bandierina di spedizione per i task di spedizione.
     * Richiede la password magazzino (mag345) ad ogni click, sia per impostare
     * che per rimuovere lo stato "spedito".  Dopo aver aggiornato lo stato,
     * salva i dati, registra la modifica per le notifiche e aggiorna il Gantt.
     * @param {Event} event L'evento click
     */
    async function handleShipStatusClick(event) {
        event.stopPropagation();
        const flag = event.target;
        const rowId = flag.dataset.rowId;
        const row = document.querySelector(`tr[data-row-id="${rowId}"]`);
        if (!row) return;
        // Richiede sempre la password magazzino
        const password = window.prompt('Inserisci la password magazzino:');
        if (password !== 'mag345') {
            window.alert('Password errata.');
            return;
        }
        const currentStatus = row.dataset.shipStatus || 'white';
        // Determina il nuovo stato: se era white diventa green (spedito), altrimenti torna white
        const newStatus = currentStatus === 'white' ? 'green' : 'white';
        row.dataset.shipStatus = newStatus;
        flag.classList.remove(`ship-status-${currentStatus}`);
        flag.classList.add(`ship-status-${newStatus}`);
        // Aggiorna il testo nel flag se necessario (opzionale, ma lasciamo sempre 'S')
        flag.textContent = 'S';
        // Raccogli i dati della riga e registra la modifica per le notifiche
        const rowData = getShippingScheduleRowData(row);
        registerShippingStatusChange(rowData, newStatus);
        // Salva localmente e sul server
        autoSaveAllData();
        if (typeof saveDataToServer === 'function') {
            try {
                await saveDataToServer();
            } catch (e) {
                console.warn('Errore nel salvataggio dei dati dopo modifica spedizione:', e);
            }
        }
        // Aggiorna il Gantt per riflettere il nuovo stato
        updateWarehouseGanttChart();
        // Controlla se ci sono avvisi da mostrare
        if (typeof checkAndNotifyShipping === 'function') {
            checkAndNotifyShipping();
        }
    }

    /**
     * Registra una modifica di stato spedizione nel localStorage per generare un avviso.
     * Ogni modifica contiene le informazioni principali della riga e un timestamp.
     * @param {Object} rowData I dati della riga di spedizione
     * @param {String} newStatus Il nuovo stato ("green" = spedito, "white" = non spedito)
     */
    function registerShippingStatusChange(rowData, newStatus) {
        try {
            let changes = [];
            try {
                changes = JSON.parse(localStorage.getItem('shippingStatusChanges') || '[]');
            } catch (e) {
                changes = [];
            }
            const id = 'SHIP-' + Date.now().toString(36) + '-' + Math.random().toString(36).substr(2, 5);
            const change = {
                id: id,
                code: rowData.codiceArticolo || rowData.codice || '',
                descrizione: rowData.descrizioneArticolo || rowData.descrizione || '',
                ov: rowData.ov || '',
                quantita: rowData.quantita || '',
                um: rowData.um || '',
                newStatus: newStatus === 'green' ? 'Spedito' : 'Non Spedito',
                timestamp: new Date().toISOString()
            };
            changes.push(change);
            localStorage.setItem('shippingStatusChanges', JSON.stringify(changes));
        } catch (e) {
            console.warn('Impossibile registrare la modifica spedizione:', e);
        }
    }

    /**
     * Controlla i cambiamenti di stato di spedizione e mostra un pop-up
     * di notifica agli utenti con permesso spedizioni.  L'utente può
     * rinviare o confermare l'avviso; in quest'ultimo caso le notifiche
     * attuali vengono marcate come riconosciute per quel livello di utente.
     */
    function checkAndNotifyShipping() {
        const perms = alertPermissions[currentUserLevel] || {};
        if (!perms.spedizioni) return;
        let changes;
        try {
            changes = JSON.parse(localStorage.getItem('shippingStatusChanges') || '[]');
        } catch (e) {
            changes = [];
        }
        if (!Array.isArray(changes) || changes.length === 0) {
            const sDiv0 = document.getElementById('shippingNotification');
            if (sDiv0) sDiv0.style.display = 'none';
            return;
        }
        let ack;
        try {
            ack = JSON.parse(localStorage.getItem('shippingAcknowledgedIds_' + currentUserLevel) || '[]');
        } catch (e) {
            ack = [];
        }
        const toShow = changes.filter(item => !ack.includes(item.id));
        if (toShow.length === 0) {
            const sDiv1 = document.getElementById('shippingNotification');
            if (sDiv1) sDiv1.style.display = 'none';
            return;
        }
        const sDiv = document.getElementById('shippingNotification');
        if (!sDiv) return;
        const pEl = sDiv.querySelector('p');
        const ulEl = sDiv.querySelector('ul');
        if (pEl) {
            pEl.textContent = 'Aggiornamenti Spedizione disponibili';
        }
        if (ulEl) {
            const itemsHtml = toShow.map(item => {
                const codeStr = item.code ? `Codice: <strong>${item.code}</strong>` : '';
                const descStr = item.descrizione ? ` - ${item.descrizione}` : '';
                const ovStr = item.ov ? `OV: <strong>${item.ov}</strong> - ` : '';
                const qtyStr = item.quantita ? ` - Quantità: ${item.quantita}${item.um ? ' ' + item.um : ''}` : '';
                return `<li>${ovStr}${codeStr}${descStr}${qtyStr} - Stato: <strong>${item.newStatus}</strong></li>`;
            }).join('');
            ulEl.innerHTML = itemsHtml;
        }
        sDiv.style.display = 'block';
        // Imposta i gestori dei pulsanti
        const postponeBtn = document.getElementById('shippingPostponeBtn');
        const okBtn = document.getElementById('shippingAcknowledgeBtn');
        const closeBtn = sDiv.querySelector('.shipping-close-btn');
        if (postponeBtn) {
            postponeBtn.onclick = () => {
                sDiv.style.display = 'none';
            };
        }
        if (okBtn) {
            okBtn.onclick = () => {
                let ackNow;
                try {
                    ackNow = JSON.parse(localStorage.getItem('shippingAcknowledgedIds_' + currentUserLevel) || '[]');
                } catch (e) {
                    ackNow = [];
                }
                toShow.forEach(item => {
                    if (!ackNow.includes(item.id)) ackNow.push(item.id);
                });
                localStorage.setItem('shippingAcknowledgedIds_' + currentUserLevel, JSON.stringify(ackNow));
                sDiv.style.display = 'none';
            };
        }
        if (closeBtn) {
            closeBtn.onclick = () => {
                sDiv.style.display = 'none';
            };
        }
    }

    /**
     * Controlla i cambiamenti di stato del magazzino e mostra un pop-up di
     * notifica agli utenti CQ.  L'utente può rinviare o confermare
     * l'avviso; in quest'ultimo caso le notifiche attuali vengono
     * marcate come riconosciute per quel livello di utente.
     */
    function checkAndNotifyWarehouse() {
        const perms = alertPermissions[currentUserLevel] || {};
        // Solo i livelli con permesso 'quarantena' ricevono la notifica magazzino.
        if (!perms.quarantena) return;
        let changes;
        try {
            changes = JSON.parse(localStorage.getItem('warehouseStatusChanges') || '[]');
        } catch (e) {
            changes = [];
        }
        if (!Array.isArray(changes) || changes.length === 0) {
            const wDiv = document.getElementById('warehouseNotification');
            if (wDiv) wDiv.style.display = 'none';
            return;
        }
        // Recupera l'elenco di ID già riconosciuti per questo livello utente
        let acknowledged;
        try {
            acknowledged = JSON.parse(localStorage.getItem('warehouseAcknowledgedIds_' + currentUserLevel) || '[]');
        } catch (e) {
            acknowledged = [];
        }
        // Filtra le notifiche non ancora riconosciute
        const filtered = changes.filter(item => !acknowledged.includes(item.id));
        if (filtered.length === 0) {
            const wDiv = document.getElementById('warehouseNotification');
            if (wDiv) wDiv.style.display = 'none';
            return;
        }
        // Raggruppa per codice per mostrare solo l'ultima notifica per codice
        const latestMap = {};
        filtered.forEach(item => {
            const key = item.code || 'UNKNOWN';
            if (!latestMap[key] || new Date(item.timestamp) > new Date(latestMap[key].timestamp)) {
                latestMap[key] = item;
            }
        });
        const toShow = Object.values(latestMap);
        const wDiv = document.getElementById('warehouseNotification');
        if (!wDiv) return;
        const pEl = wDiv.querySelector('p');
        const ulEl = wDiv.querySelector('ul');
        if (pEl) {
            pEl.textContent = 'Aggiornamenti Magazzino disponibili';
        }
        if (ulEl) {
            const itemsHtml = toShow.map(item => {
                const codeStr = item.code ? `Codice: <strong>${item.code}</strong>` : '';
                const descStr = item.descrizione ? ` - ${item.descrizione}` : '';
                const ovStr = item.ov ? `OV: <strong>${item.ov}</strong> - ` : '';
                const qtyStr = item.quantita ? ` - Quantità: ${item.quantita}${item.um ? ' ' + item.um : ''}` : '';
                return `<li>${ovStr}${codeStr}${descStr}${qtyStr} - Stato: <strong>${item.newStatus}</strong></li>`;
            }).join('');
            ulEl.innerHTML = itemsHtml;
        }
        wDiv.style.display = 'block';
        // Imposta i gestori dei pulsanti
        const postponeBtn = document.getElementById('warehousePostponeBtn');
        const okBtn = document.getElementById('warehouseAcknowledgeBtn');
        const closeBtn = wDiv.querySelector('.warehouse-close-btn');
        if (postponeBtn) {
            postponeBtn.onclick = () => {
                wDiv.style.display = 'none';
            };
        }
        if (okBtn) {
            okBtn.onclick = () => {
                // Aggiorna l'elenco degli ID riconosciuti per questo utente
                let ack = [];
                try {
                    ack = JSON.parse(localStorage.getItem('warehouseAcknowledgedIds_' + currentUserLevel) || '[]');
                } catch (e) {
                    ack = [];
                }
                toShow.forEach(item => {
                    if (!ack.includes(item.id)) ack.push(item.id);
                });
                localStorage.setItem('warehouseAcknowledgedIds_' + currentUserLevel, JSON.stringify(ack));
                wDiv.style.display = 'none';
            };
        }
        if (closeBtn) {
            closeBtn.onclick = () => {
                wDiv.style.display = 'none';
            };
        }
    }
    /**
     * Controlla i cambiamenti di stato CQ/QA e notifica l'utente se necessario.
     * Vengono mostrati gli avvisi in un pop-up simile a quello ADR.
     * L'utente può rinviare o confermare l'avviso, in quest'ultimo caso
     * l'elenco attuale viene marcato come già visualizzato nel localStorage.
     */
    function checkAndNotifyQuality() {
        // Verifica i permessi per CQ e QA per il livello corrente
        const perms = alertPermissions[currentUserLevel] || {};
        if (!perms.CQ && !perms.QA) return;
        let changes;
        try {
            changes = JSON.parse(localStorage.getItem('qualityStatusChanges') || '[]');
        } catch (e) {
            changes = [];
        }
        if (!Array.isArray(changes) || changes.length === 0) {
            // Nessuna modifica di qualità: assicurati che l'avviso CQ/QA sia nascosto
            const qDiv = document.getElementById('qualityNotification');
            if (qDiv) qDiv.style.display = 'none';
            return;
        }
        // Recupera la lista di ID già accettati per questo livello
        let acknowledged = [];
        try {
            acknowledged = JSON.parse(localStorage.getItem('qualityAcknowledgedIds_' + currentUserLevel) || '[]');
        } catch (e) {
            acknowledged = [];
        }
        // Filtra le modifiche in base ai permessi dell'utente e non già riconosciute
        const filtered = changes.filter(item => {
            const permCheck = (item.type === 'CQ' && perms.CQ) || (item.type === 'QA' && perms.QA);
            const notAck = !acknowledged.includes(item.id);
            return permCheck && notAck;
        });
        if (filtered.length === 0) {
            // Tutte le modifiche sono state riconosciute o non sono pertinenti: nascondi l'avviso
            const qDiv = document.getElementById('qualityNotification');
            if (qDiv) qDiv.style.display = 'none';
            return;
        }
        // Raggruppa per combinazione di codice/lotto/type e mantieni solo l'ultima versione
        const latestMap = {};
        filtered.forEach(item => {
            const key = `${item.type}-${item.code}-${item.lotto}`;
            if (!latestMap[key] || new Date(item.timestamp) > new Date(latestMap[key].timestamp)) {
                latestMap[key] = item;
            }
        });
        const toShow = Object.values(latestMap);
        // Costruisci HTML degli elementi della lista: includi tipo, OV, OP, codice, descrizione, lotto, quantità e stato
        const itemsHtml = toShow.map(item => {
            const typeLabel = item.type === 'CQ' ? 'CQ' : 'QA';
            const ovStr = item.ov ? `OV: <strong>${item.ov}</strong> - ` : '';
            const opStr = item.op ? `OP: <strong>${item.op}</strong> - ` : '';
            const descStr = item.descrizione ? ` - ${item.descrizione}` : '';
            // Mostra sempre il lotto (anche se vuoto), valorizzando con N/D se assente
            const lottoVal = item.lotto ? item.lotto : 'N/D';
            const lottoStr = ` - Lotto: <strong>${lottoVal}</strong>`;
            let qtyStr = '';
            if (item.quantita) {
                qtyStr = ` - Quantità: ${item.quantita}${item.um ? ' ' + item.um : ''}`;
            }
            return `<li><strong>${typeLabel}</strong> - ${ovStr}${opStr}Codice: <strong>${item.code}</strong>${descStr}${lottoStr}${qtyStr} - Stato: <strong>${item.newStatus}</strong></li>`;
        }).join('');
        const qDiv = document.getElementById('qualityNotification');
        if (!qDiv) return;
        const pEl = qDiv.querySelector('p');
        const ulEl = qDiv.querySelector('ul');
        if (pEl) {
            const hasCQ = toShow.some(item => item.type === 'CQ');
            const hasQA = toShow.some(item => item.type === 'QA');
            let msg;
            if (hasCQ && hasQA) msg = 'Aggiornamenti CQ e QA disponibili';
            else if (hasCQ) msg = 'Aggiornamenti CQ disponibili';
            else msg = 'Aggiornamenti QA disponibili';
            pEl.textContent = msg;
        }
        if (ulEl) ulEl.innerHTML = itemsHtml;
        qDiv.style.display = 'block';
        const postponeBtn = document.getElementById('qualityPostponeBtn');
        const okBtn = document.getElementById('qualityAcknowledgeBtn');
        if (postponeBtn) {
            postponeBtn.style.display = 'inline-flex';
            postponeBtn.disabled = false;
            postponeBtn.style.pointerEvents = 'auto';
            postponeBtn.onclick = () => {
                qDiv.style.display = 'none';
            };
        }
        if (okBtn) {
            okBtn.style.display = 'inline-flex';
            okBtn.disabled = false;
            okBtn.style.pointerEvents = 'auto';
            okBtn.onclick = () => {
                // Aggiungi tutti gli id mostrati alla lista di riconoscimenti per questo livello
                const ackList = new Set(acknowledged);
                toShow.forEach(item => ackList.add(item.id));
                localStorage.setItem('qualityAcknowledgedIds_' + currentUserLevel, JSON.stringify(Array.from(ackList)));
                qDiv.style.display = 'none';
            };
        }
        const closeBtn = qDiv.querySelector('.quality-close-btn');
        if (closeBtn) {
            closeBtn.style.display = 'inline';
            closeBtn.style.pointerEvents = 'auto';
            closeBtn.onclick = () => {
                qDiv.style.display = 'none';
            };
        }
    }

/**
 * Apre l'editor per scrivere/modificare i commenti dopo l'inserimento della password.
 * VERSIONE CORRETTA PER AGGIORNAMENTO IMMEDIATO DEL TOOLTIP.
 */
/**
 * VERSIONE DEFINITIVA - Combina salvataggio corretto e aggiornamento visivo istantaneo.
 * Apre l'editor dei commenti QA, salva la modifica e aggiorna immediatamente la UI.
 */


    // --- INIZIALIZZAZIONE E ATTIVAZIONE ---

    // Sostituisce il vecchio caricamento da localStorage con quello dal server
    loadDataFromServer();

    // Crea una versione debounced di saveDataToServer per evitare salvataggi troppo frequenti.
    // Viene eseguita al massimo una volta ogni 800 ms.
    const debouncedSave = (typeof debounce === 'function') ? debounce(() => saveDataToServer(), 800) : (async () => saveDataToServer());
    
    // Attiva il salvataggio automatico dopo ogni modifica.
    // Aggiungi questo event listener a tutte le tabelle modificabili.
    // Include un controllo per evitare salvataggi e alert prima del login.
    document.querySelector('.container').addEventListener('change', (event) => {
        // Considera solo le modifiche su input, select o textarea
        if (!event.target.matches('input, select, textarea')) return;
        // Recupera l'overlay del login e lo stato di login
        const loginOverlayEl = document.getElementById('loginOverlay');
        const overlayVisible = loginOverlayEl && loginOverlayEl.style.display !== 'none';
        // Se l'overlay è visibile o non siamo loggati, non salvare e non generare alert
        if (overlayVisible || !currentUserLevel || currentUserLevel <= 0) {
            return;
        }
        // Usa la funzione debounced per ridurre i salvataggi a raffica
        debouncedSave();
    });

    // Imposta un intervallo per ricaricare i dati periodicamente. Qui abbiamo
    // scelto un intervallo di 5 minuti (300.000 ms) per sincronizzare
    // automaticamente le modifiche degli altri utenti. Se desideri un
    // intervallo diverso, modifica il valore in millisecondi.
    setInterval(loadDataFromServer, 300000);


    duplicateRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('.row-selector:checked');
        if (selectedRows.length === 0) {
            await showAlert('Seleziona almeno una riga da duplicare.');
            return;
        }

        const confirmed = await showConfirm(`Sei sicuro di voler duplicare ${selectedRows.length} riga/e selezionata/e?`);
        if (confirmed) {
            selectedRows.forEach(checkbox => {
                const originalRow = checkbox.closest('tr');
                const rowData = getRowData(originalRow);
                const newRow = createRow(rowData);
                originalRow.after(newRow);
                validateRow(newRow);
            });
            await showAlert('Riga/e duplicata/e con successo.');
            updateGanttChart();
            updateWarehouseGanttChart();
            updateDailyProductionTable();
            updateAnalisiTable();
            runFullCheck();
            addLogEntry(`Duplicate ${selectedRows.length} riga/e.`);
            autoSaveAllData();
        }
    });

    function getRowData(row) {
        const codeValue = row.querySelector('.code-input').value;
        const confezionamentoPezzi = row.querySelector('.packaging-pieces-input').value;
        const confezionamentoKgPerPiece = row.querySelector('.packaging-kg-per-piece-input').value;
        const confezionamentoUnit = row.querySelector('.col-confez-kg-pezzo .unit-select').value;

        let rawConfezionamentoDettaglio = '';
        if (isMedicalDeviceCode(codeValue)) {
            if (confezionamentoPezzi) {
                rawConfezionamentoDettaglio = `${confezionamentoPezzi}X${confezionamentoKgPerPiece}mL`;
            }
        } else {
            if (confezionamentoPezzi && confezionamentoKgPerPiece && confezionamentoUnit) {
                rawConfezionamentoDettaglio = `${confezionamentoPezzi}X${confezionamentoKgPerPiece}${confezionamentoUnit}`;
            } else if (confezionamentoPezzi && confezionamentoUnit) {
                rawConfezionamentoDettaglio = `${confezionamentoPezzi}${confezionamentoUnit}`;
            } else if (confezionamentoPezzi) {
                rawConfezionamentoDettaglio = `${confezionamentoPezzi}`;
            }
        }

        return {
            codice: row.querySelector('.code-input').value,
            prodotto: row.querySelector('.product-input').value,
            cliente: row.querySelector('.client-input').value,
            quantitaRichiesta: row.querySelector('.qty-requested-input').value,
            quantitaRichiestaUnit: row.querySelector('.qty-requested-unit-select').value,
            giacenzaMagazzino: row.querySelector('.stock-input').value,
            quantitaDaProdurre: row.querySelector('.qty-to-produce-input').value,
            materiePrime: row.querySelector('.materie-prime-select').value,
            macchinari: row.querySelector('.machine-input').value,
            operatore: row.querySelector('.operator-input').value,
            confezionamentoPezzi: confezionamentoPezzi,
            confezionamentoKgPerPiece: confezionamentoKgPerPiece,
            confezionamentoUnit: confezionamentoUnit,
            rawConfezionamentoDettaglio: rawConfezionamentoDettaglio,
            produzioneData: row.querySelector('.production-date-input').value,
            giorniDiProduzione: row.querySelector('.production-days-input').value,
            dataConfezionamento: row.querySelector('.packaging-date-input').value,
            codiceConfezionamento: row.querySelector('.packaging-code-input').value,
            lottoSC: row.querySelector('.lotto-sc-input').value,
            materialeConfezionamento: row.querySelector('.materiale-confez-select').value,
            dataSpedizione: row.querySelector('.shipping-date-input').value,
            note: row.querySelector('.notes-input').value
        };
    }


// ========================================================================
    // ==> FUNZIONI NUOVE PER LA TABELLA MEDICAL DEVICE <==
    // ========================================================================

    /**
     * Crea una riga HTML per la tabella dei dispositivi medici.
     */
    function createMedicalDeviceRow(rowData = {}, isManual = false) {
        const row = document.createElement('tr');
        
        let siringhePerScatola = '';
        const confezionamentoString = rowData.rawConfezionamentoDettaglio || 
                                      (rowData.confezionamentoPezzi ? `${rowData.confezionamentoPezzi}x` : '');

        const match = confezionamentoString.match(/^(\d)x/);
        if (match && (match[1] === '1' || match[1] === '2')) {
            siringhePerScatola = match[1];
        }
        
        const volumeProduzione = rowData.quantitaDaProdurre ? `${rowData.quantitaDaProdurre} Kg` : '';

        const isEditable = (currentUserLevel === 3 || currentUserLevel === 6);
        const readOnlyAttr = isEditable ? '' : 'readonly';

        row.innerHTML = `
            <td><input type="text" value="${rowData.codice || ''}" ${isManual ? '' : 'readonly'}></td>
            <td><input type="text" value="${rowData.prodotto || ''}" ${isManual ? '' : 'readonly'}></td>
            <td><input type="text" value="${rowData.cliente || ''}" ${isManual ? '' : 'readonly'}></td>
            <td><input type="text" value="${rowData.confezionamentoPezzi || ''}" ${isManual ? '' : 'readonly'}></td>
            <td><input type="text" class="scarti-input" value="${rowData.scarti || ''}" ${readOnlyAttr}></td>
            <td><input type="text" value="${volumeProduzione}" ${isManual ? '' : 'readonly'}></td>
            <td><input type="text" value="${siringhePerScatola}" ${isManual ? '' : 'readonly'}></td>
        `;
        
        if(isManual){
             row.querySelectorAll('input').forEach(input => {
                if(!input.classList.contains('scarti-input')) {
                    input.readOnly = false;
                }
             });
        }
        
        return row;
    }

    /**
     * Aggiorna e filtra la tabella dei dispositivi medici.
     */
    function updateMedicalDeviceProductionTable() {
        if (!medicalDeviceTableBody) return;

        // Se esistono dati di produzione medicale salvati localmente, usali direttamente
        // per popolare la tabella e ritorna.  Questo previene la sovrascrittura dei
        // dati importati con i valori della tabella principale e garantisce che le
        // ultime righe importate restino visibili finché non viene effettuato un nuovo
        // import manuale.  I filtri e le date verranno applicati successivamente solo
        // se non ci sono dati importati in memoria.
        try {
            const storedProd = localStorage.getItem('medicalProductionData');
            if (storedProd) {
                const parsed = JSON.parse(storedProd);
                populateMedicalDeviceProductionTable(parsed);
                return;
            }
        } catch (e) {
            console.warn('Errore nel caricamento della produzione medicale:', e);
        }

        // Se esistono dati di produzione medicale importati, usali per popolare la tabella
        try {
            const storedProd = localStorage.getItem('medicalProductionData');
            if (storedProd) {
                const parsed = JSON.parse(storedProd);
                populateMedicalDeviceProductionTable(parsed);
                return;
            }
        } catch (e) {
            console.warn('Errore nel caricamento della produzione medicale:', e);
        }
        
        const scartiValues = new Map();
        medicalDeviceTableBody.querySelectorAll('tr').forEach(row => {
            const codice = row.cells[0].querySelector('input').value;
            const scarti = row.cells[4].querySelector('input').value;
            if (codice && scarti) {
                scartiValues.set(codice, scarti);
            }
        });

        medicalDeviceTableBody.innerHTML = '';
        const allMainTableData = getAllTableData();
        const medicalDevicesData = allMainTableData.filter(row => isMedicalDeviceCode(row.codice));

        const startDate = medicalDeviceStartDateInput._flatpickr.selectedDates[0];
        const endDate = medicalDeviceEndDateInput._flatpickr.selectedDates[0];
        if (endDate) endDate.setHours(23, 59, 59, 999);

        const filterCodiceText = filterMedicalDeviceCodice.value.toLowerCase();
        const filterDescrizioneText = filterMedicalDeviceDescrizione.value.toLowerCase();
        const filterClienteText = filterMedicalDeviceCliente.value.toLowerCase();

        const filteredMedicalData = medicalDevicesData.filter(row => {
            const prodDateParts = row.produzioneData.split('/');
            if (prodDateParts.length !== 3) return false;
            const rowDate = new Date(parseInt(prodDateParts[2]), parseInt(prodDateParts[1]) - 1, parseInt(prodDateParts[0]));
            const dateMatch = (!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate);
            if (!dateMatch) return false;

            const codiceMatch = !filterCodiceText || (row.codice || '').toLowerCase().includes(filterCodiceText);
            const descrizioneMatch = !filterDescrizioneText || (row.prodotto || '').toLowerCase().includes(filterDescrizioneText);
            const clienteMatch = !filterClienteText || (row.cliente || '').toLowerCase().includes(filterClienteText);

            return codiceMatch && descrizioneMatch && clienteMatch;
        });

        filteredMedicalData.forEach(rowData => {
            if (scartiValues.has(rowData.codice)) {
                rowData.scarti = scartiValues.get(rowData.codice);
            }
            medicalDeviceTableBody.appendChild(createMedicalDeviceRow(rowData));
        });

        // Rende nuovamente ordinabile la tabella dispositivi medici dopo l'aggiornamento
        if (typeof makeTableSortable === 'function') {
            makeTableSortable(document.getElementById('medicalDeviceProductionTable'));
        }
    }

    deleteRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('.row-selector:checked');
        if (selectedRows.length === 0) {
            await showAlert('Seleziona almeno una riga da eliminare.');
            return;
        }

        const confirmed = await showConfirm(`Sei sicuro di voler eliminare ${selectedRows.length} riga/e selezionata/e?`);
        if (confirmed) {
            selectedRows.forEach(checkbox => {
                checkbox.closest('tr').remove();
            });
            await showAlert('Riga/e eliminata/e con successo.');
            updateGanttChart();
            updateWarehouseGanttChart();
            updateDailyProductionTable();
            updateAnalisiTable();
            runFullCheck();
            addLogEntry(`Eliminate ${selectedRows.length} riga/e.`);
            autoSaveAllData();
        }
    });

    function getAllTableData() {
        const data = [];
        productionTableBody.querySelectorAll('tr').forEach(row => {
            data.push(getRowData(row));
        });
        return data;
    }

    function getAllSalesOrderData() {
        const data = [];
        salesOrderTableBody.querySelectorAll('tr').forEach(row => {
            data.push(getSalesOrderRowData(row));
        });
        return data;
    }

/**
 * NUOVA FUNZIONE DI SUPPORTO
 * Crea un ID stabile e univoco per una riga di spedizione basato sui suoi dati.
 */
function getStableRowId(rowData) {
    const ov = String(rowData.ov || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    const code = String(rowData.codiceArticolo || '').trim().replace(/[^a-zA-Z0-9]/g, '');
    return `shipping-stable-${ov}-${code}`;
}

function loadAllAutoSavedData() {
    // 1. CARICA I DATI DEL PROGRAMMA DI PRODUZIONE
    const savedProductionData = localStorage.getItem('production_data_autosave');
    if (savedProductionData) {
        try {
            const parsedData = JSON.parse(savedProductionData);
            productionTableBody.innerHTML = '';
            parsedData.forEach(rowData => {
                productionTableBody.appendChild(createRow(rowData));
            });
            console.log('Dati di produzione salvati automaticamente caricati.');
        } catch (e) {
            console.error('Errore durante il caricamento dei dati di produzione salvati automaticamente:', e);
        }
    }

    // 2. CARICA I DATI DEGLI ORDINI DI VENDITA
    const savedSalesOrderData = localStorage.getItem('sales_order_data_autosave');
    if (savedSalesOrderData) {
        try {
            const parsedData = JSON.parse(savedSalesOrderData);
            salesOrderTableBody.innerHTML = '';
            parsedData.forEach(rowData => {
                salesOrderTableBody.appendChild(createSalesOrderRow(rowData));
            });
            console.log('Dati degli ordini di vendita salvati automaticamente caricati.');
        } catch (e) {
            console.error('Errore durante il caricamento dei dati degli ordini di vendita salvati automaticamente:', e);
        }
    }
    
    // 3. CARICA I DATI DEL PROGRAMMA DI SPEDIZIONE
    const savedShippingData = localStorage.getItem('shipping_schedule_data_autosave');
    if (savedShippingData) {
        try {
            const parsedData = JSON.parse(savedShippingData);
            shippingScheduleTableBody.innerHTML = '';
            parsedData.forEach(rowData => {
                shippingScheduleTableBody.appendChild(createShippingScheduleRow(rowData));
            });
            console.log('Dati del programma di spedizione salvati automaticamente caricati.');
        } catch (e) {
            console.error('Errore durante il caricamento dei dati del programma di spedizione:', e);
        }
    }

    // 4. CARICA I DATI DEL PROGRAMMA DI ARRIVO
    const savedArrivalData = localStorage.getItem('arrival_schedule_data_autosave');
    if (savedArrivalData) {
        try {
            const parsedData = JSON.parse(savedArrivalData);
            if (arrivalScheduleTableBody) arrivalScheduleTableBody.innerHTML = '';
            parsedData.forEach(rowData => {
                if (arrivalScheduleTableBody) arrivalScheduleTableBody.appendChild(createArrivalScheduleRow(rowData));
            });
            console.log('Dati del programma di arrivo salvati automaticamente caricati.');
        } catch (e) {
            console.error('Errore durante il caricamento dei dati del programma di arrivo:', e);
        }
    }
    
    // ===== BLOCCO AGGIUNTO PER CARICARE LA MERCE NON ARRIVATA =====
    const savedOverdueArrivalData = localStorage.getItem('overdue_arrival_data_autosave');
    if (savedOverdueArrivalData) {
        try {
            const parsedData = JSON.parse(savedOverdueArrivalData);
            populateOverdueTable(parsedData);
            console.log('Dati della merce non arrivata salvati automaticamente caricati.');
        } catch (e) {
            console.error('Errore durante il caricamento dei dati della merce non arrivata:', e);
        }
    }
    // =================================================================
}

    

    loadDataBtn.addEventListener('click', async () => {
        const savedKeys = Object.keys(localStorage).filter(key => key.startsWith('production_data_') && key !== 'production_data_autosave');
        const savedNames = savedKeys.map(key => key.replace('production_data_', ''));

        if (savedNames.length === 0) {
            await showAlert('Nessun dato salvato trovato.');
            return;
        }

        const selectedAction = await showSelectionModal('Carica Dati', 'Seleziona un salvataggio:', savedNames);

        if (selectedAction === null) {
            await showAlert('Operazione di caricamento annullata.');
            return;
        }

        if (selectedAction === 'delete') {
            const selectedSaveName = document.getElementById('modalSelect').value;
            const confirmedDelete = await showConfirm(`Sei sicuro di voler eliminare il salvataggio "${selectedSaveName}"?`);
            if (confirmedDelete) {
                localStorage.removeItem(`production_data_${selectedSaveName}`);
                addLogEntry(`Salvataggio dati produzione "${selectedSaveName}" eliminato.`);
                await showAlert(`Salvataggio "${selectedSaveName}" eliminato.`);
                const remainingKeys = Object.keys(localStorage).filter(key => key.startsWith('production_data_'));
                if (remainingKeys.length > 0) {
                    loadDataBtn.click();
                } else {
                    productionTableBody.innerHTML = '';
                }
            }
            return;
        }

        const selectedSaveName = selectedAction;
        try {
            const key = `production_data_${selectedSaveName}`;
            const savedData = localStorage.getItem(key);
            if (savedData) {
                const parsedData = JSON.parse(savedData);
                productionTableBody.innerHTML = '';
                parsedData.forEach(rowData => {
                    productionTableBody.appendChild(createRow(rowData));
                });
                await showAlert(`Dati caricati con successo da "${selectedSaveName}".`);
                updateGanttChart();
                updateWarehouseGanttChart();
                updateScrollButtons();
                applyFilter();
                updateDailyProductionTable();
                updateAnalisiTable();
                runFullCheck();
                addLogEntry(`Dati di produzione caricati manualmente: "${selectedSaveName}".`);
                autoSaveAllData();
            } else {
                await showAlert('Salvataggio non trovato.');
            }
        } catch (e) {
            console.error("Errore durante il caricamento dei dati:", e);
            addLogEntry(`Errore caricamento manuale dati produzione: ${e.message}.`);
            await showAlert(`Errore durante il caricamento dei dati: ${e.message}.`);
        }
    });

    sendEmailBtn.addEventListener('click', () => {
        const recipient = 'rossella.crippa@iralab.it';
        const subject = encodeURIComponent('Aggiornamento Programma di Produzione');

        const tableRowsData = getAllTableData();
        const currentDateTime = new Date().toLocaleString('it-IT');

        let body = `Gentile Rossella,\n\n`;
        body += `Ti invio un aggiornamento sul programma di produzione.\n`;
        body += `\nData e Ora dell'aggiornamento: ${currentDateTime}\n`;
        body += `Numero totale di attività nel programma: ${tableRowsData.length}\n\n`;
        body += `Dettagli delle attività:\n`;

        tableRowsData.forEach((row, index) => {
            body += `\n--- Attività ${index + 1} ---\n`;
            body += `Codice: ${row.codice || 'N/D'}\n`;
            body += `Prodotto: ${row.prodotto || 'N/D'}\n`;
            body += `Cliente: ${row.cliente || 'N/D'}\n`;
            body += `Quantità Richiesta: ${row.quantitaRichiesta || 'N/D'} ${row.quantitaRichiestaUnit || ''}\n`;
            body += `Giacenza Magazzino: ${row.giacenzaMagazzino || 'N/D'} Kg\n`;
            body += `Quantità da Produrre: ${row.quantitaDaProdurre || 'N/D'} Kg\n`;
            body += `Materie Prime: ${row.materiePrime || 'N/D'}\n`;
            body += `Macchinari: ${row.macchinari || 'N/D'}\n`;
            body += `Operatore: ${row.operatore || 'N/D'}\n`;
            body += `Confezionamento Richiesto (Pezzi): ${row.confezionamentoPezzi || 'N/D'}\n`;
            body += `Confezionamento Richiesto (Kg/Pezzo): ${row.confezionamentoKgPerPiece || 'N/D'} ${row.confezionamentoUnit || ''}\n`;
            body += `Data di Produzione: ${row.produzioneData || 'N/D'}\n`;
            body += `Giorni di Produzione: ${row.giorniDiProduzione || 'N/D'}\n`;
            body += `Data di Confezionamento: ${row.dataConfezionamento || 'N/D'}\n`;
            body += `Codice Confezionamento: ${row.codiceConfezionamento || 'N/D'}\n`;
            body += `Lotto SC: ${row.lottoSC || 'N/D'}\n`;
            body += `Materiale Confezionamento: ${row.materialeConfezionamento || 'N/D'}\n`;
            body += `Data di Spedizione: ${row.dataSpedizione || 'N/D'}\n`;
            body += `Note: ${row.note || 'N/D'}\n`;
        });

        body += `\nPotrai trovare i dettagli completi nel file allegato o consultando l'applicazione.\n\n`;
        body += `Cordiali saluti,\nIl tuo nome/La tua azienda`;

        const mailtoLink = `mailto:${recipient}?subject=${subject}&body=${encodeURIComponent(body)}`;
        window.location.href = mailtoLink;
        addLogEntry(`Email programma di produzione inviata a ${recipient}.`);
    });


    exportDataBtn.addEventListener('click', async () => {
        const data = getAllTableData();
        if (data.length === 0) {
            await showAlert('Nessun dato da esportare.');
            return;
        }

        const headers = [
            "Codice", "Prodotto", "Cliente", "Quantità Richiesta", "Unità Quantità Richiesta", "Giacenza Magazzino (Kg)", "Quantità da Produrre (Kg)",
            "Materie Prime", "Macchinari", "Operatore", "Confezionamento Richiesto (Numero Pezzi)", "Confezionamento Richiesto (Kg/Pezzo)", "Unità Confezionamento",
            "Data di Produzione", "Giorni di Produzione", "Data di Confezionamento", "Codice Confezionamento", "Lotto SC", "Materiale Confezionamento", "Data di Spedizione", "Note"
        ];

        let csvContent = headers.map(h => `"${h.replace(/"/g, '""')}"`).join(';') + '\n';

        data.forEach(row => {
            const rowValues = [
                row.codice, row.prodotto, row.cliente, row.quantitaRichiesta, row.quantitaRichiestaUnit, row.giacenzaMagazzino, row.quantitaDaProdurre,
                row.materiePrime, row.macchinari, row.operatore, row.confezionamentoPezzi, row.confezionamentoKgPerPiece, row.confezionamentoUnit,
                row.produzioneData, row.giorniDiProduzione, row.dataConfezionamento, row.codiceConfezionamento, row.lottoSC, row.materialeConfezionamento, row.dataSpedizione, row.note
            ];
            csvContent += rowValues.map(val => `"${String(val).replace(/"/g, '""')}"`).join(';') + '\n';
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        const todayForFileName = new Date().toLocaleDateString('it-IT').replace(/\//g, '.');

        let fileNameSuffix = '';
        const filterCol1 = filterColumn1Select.value;
        const filterVal1 = filterValue1Input.value.trim();
        const filterCol2 = filterColumn2Select.value;
        const filterVal2 = filterValue2Input.value.trim();

        if (filterCol1 && filterVal1) {
            fileNameSuffix += `_filtro1-${filterCol1}-${filterVal1.replace(/[^a-zA-Z0-9]/g, '')}`;
        }
        if (filterCol2 && filterVal2) {
            fileNameSuffix += `_filtro2-${filterCol2}-${filterVal2.replace(/[^a-zA-Z0-9]/g, '')}`;
        }

        link.download = `programma_produzione_${todayForFileName}${fileNameSuffix}.csv`;

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        addLogEntry(`Dati produzione esportati come CSV: "${link.download}".`);
    });
    async function compareAndApplyPPChanges(newData, source) {
    const oldData = getAllTableData();
    const oldDataMap = new Map(oldData.map(row => [row.codice, row]));
    const newDataMap = new Map(newData.map(row => [row.codice, row]));
    let changesFound = false;
    let logHeaderAdded = false;

    const addLogHeader = () => {
        if (!logHeaderAdded) {
            addLogEntry(`--- Inizio confronto importazione PP: ${source} ---`);
            logHeaderAdded = true;
        }
    };

    for (const [codice, newRowData] of newDataMap.entries()) {
        const oldRowData = oldDataMap.get(codice);
        const logIdentifier = `riga (Codice: ${newRowData.codice}, Prodotto: ${newRowData.prodotto}${newRowData.lottoSC ? ', Lotto: ' + newRowData.lottoSC : ''})`;

        if (!oldRowData) {
            addLogHeader();
            addLogEntry(`Aggiunta ${logIdentifier}.`);
            productionTableBody.appendChild(createRow(newRowData));
            changesFound = true;
        } else {
            const fieldsToCompare = ['prodotto', 'cliente', 'quantitaRichiesta', 'giacenzaMagazzino', 'quantitaDaProdurre', 'produzioneData', 'dataConfezionamento', 'dataSpedizione', 'lottoSC', 'note', 'rawConfezionamentoDettaglio'];
            let modifications = [];

            fieldsToCompare.forEach(field => {
                let oldValue = String(oldRowData[field] || "").trim();
                let newValue = String(newRowData[field] || "").trim();
                if (field === 'rawConfezionamentoDettaglio') {
                    oldValue = oldValue.toLowerCase().replace(',', '.');
                    newValue = newValue.toLowerCase().replace(',', '.');
                }
                if (oldValue !== newValue) {
                    modifications.push(`campo '${field}' da '${oldRowData[field] || ""}' a '${newRowData[field] || ""}'`);
                }
            });

            if (modifications.length > 0) {
                addLogHeader();
                addLogEntry(`Modificata ${logIdentifier}: ${modifications.join(', ')}.`);
                const rowElement = Array.from(productionTableBody.querySelectorAll('.code-input')).find(input => input.value === codice)?.closest('tr');
                if (rowElement) {
                    const updatedRow = createRow(newRowData);
                    rowElement.replaceWith(updatedRow);
                }
                changesFound = true;
            }
            oldDataMap.delete(codice);
        }
    }

    for (const [codice, oldRowData] of oldDataMap.entries()) {
        addLogHeader();
        const logIdentifier = `riga (Codice: ${oldRowData.codice}, Prodotto: ${oldRowData.prodotto}${oldRowData.lottoSC ? ', Lotto: ' + oldRowData.lottoSC : ''})`;
        addLogEntry(`Tolta ${logIdentifier} perché non presente nel nuovo file.`);
        const rowElement = Array.from(productionTableBody.querySelectorAll('.code-input')).find(input => input.value === codice)?.closest('tr');
        if (rowElement) rowElement.remove();
        changesFound = true;
    }

    if (changesFound) {
        addLogEntry(`--- Fine confronto: Modifiche applicate con successo. ---`);
        productionTableBody.querySelectorAll('tr').forEach(row => validateRow(row));
        
        // ===================================================================
        // ==> CHIAMATA ALLA FUNZIONE DI AGGIORNAMENTO GLOBALE <==
        // ===================================================================
        updateAllUIComponents(); // Questa chiamata risolve il problema di sincronizzazione
        
        autoSaveAllData();
        await showAlert('Importazione completata. Sono state trovate e applicate delle modifiche. Controlla il logbook per i dettagli.');
    } else {
        await showAlert('Importazione completata. Nessuna modifica trovata rispetto ai dati attuali.');
    }
}


async function processPPFile(file) {
        // --- INIZIO DELLA MODIFICA FONDAMENTALE ---
        console.log("Avvio processo importazione PP. Forzo il refresh dei dati statici...");
        addLogEntry(`Avvio importazione PP. Ricarico Referenze e Piano Analitico dalla memoria...`);

        // 1. RICARICA OBBLIGATORIAMENTE I DATI STATICI PRIMA DI FARE QUALSIASI ALTRA COSA
        refreshStaticDataFromStorage();

        // --- FINE DELLA MODIFICA FONDAMENTALE ---

        const importedData = [];
        let headerRowIndex = -1;
        let skipNextRow = false;

        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });

            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

                skipNextRow = false;
                headerRowIndex = -1;
                for (let i = 0; i < json.length; i++) {
                    if (Array.isArray(json[i]) &&
                        String(json[i][0] || '').trim().toLowerCase() === 'codice' &&
                        String(json[i][8] || '').trim().toLowerCase() === 'produz.' &&
                        String(json[i][9] || '').trim().toLowerCase() === 'confez.') {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (headerRowIndex === -1) {
                    console.warn(`The sheet "${sheetName}" does not contain the expected header row. Skipping.`);
                    return;
                }

                const colIndex = {
                    'CODICE': 0, 'PRODOTTO': 1, 'CLIENTE': 3, 'QTA_RICHIESTA': 4,
                    'QTA_DA_PRODURRE': 5, 'CONFEZIONAMENTO_DETTAGLIO': 6, 'GIACENZA_MAGAZZINO': 7,
                    'DATA_PRODUZIONE': 8, 'DATA_CONFEZIONAMENTO': 9, 'DATA_SPEDIZIONE': 10,
                    'LOTTO_SC': 11, 'NOTE': 12
                };

                const dataRows = json.slice(headerRowIndex + 1);

                dataRows.forEach((excelRow, rowIndex) => {
                    if (skipNextRow) {
                        skipNextRow = false;
                        return;
                    }

                    if (excelRow.every(cell => cell === null || cell === undefined || String(cell).trim() === '')) {
                        return;
                    }

                    const getVal = (key) => excelRow[colIndex[key]] !== undefined ? excelRow[colIndex[key]] : '';
                    const rawCodice = getVal('CODICE');
                    const rawProdotto = getVal('PRODOTTO');

                    if (!rawCodice || !rawProdotto) {
                        return;
                    }

                    const isMedical = isMedicalDeviceCode(rawCodice);
                    let packagingDetails;

                    if (isMedical) {
                        const mainRowPackagingDetail = getVal('CONFEZIONAMENTO_DETTAGLIO');
                        const mainRowParsed = parsePackagingString(mainRowPackagingDetail);

                        let pezzi = mainRowParsed.pezzi;
                        let kgPerPezzo = '';
                        let unit = 'mL';

                        if (rowIndex + 1 < dataRows.length) {
                            const nextExcelRow = dataRows[rowIndex + 1];
                            const nextRowCode = nextExcelRow[colIndex['CODICE']];
                            if (!nextRowCode || String(nextRowCode).trim() === '') {
                                const nextRowPackagingDetail = String(nextExcelRow[colIndex['CONFEZIONAMENTO_DETTAGLIO']] || '').trim();
                                if (nextRowPackagingDetail) {
                                    const medicalMatch = nextRowPackagingDetail.match(/(?:\d+[Xx])?(\d+([.,]\d+)?)/);
                                    kgPerPezzo = medicalMatch ? medicalMatch[1].replace(',', '.') : '';

                                    const unitMatch = nextRowPackagingDetail.match(/[a-zA-Z]+$/);
                                    if (unitMatch) {
                                        unit = normalizeUnit(unitMatch[0]);
                                    }
                                    skipNextRow = true;
                                }
                            }
                        }

                        packagingDetails = {
                            pezzi: pezzi,
                            kgPerPezzo: kgPerPezzo,
                            unit: unit
                        };
                    } else {
                        const rawConfezionamentoDettaglio = getVal('CONFEZIONAMENTO_DETTAGLIO');
                        packagingDetails = parsePackagingString(rawConfezionamentoDettaglio);
                    }

                    const qtyRequested = parseNumericValue(getVal('QTA_RICHIESTA'));
                    const qtyRequestedUnit = normalizeUnit(String(getVal('QTA_RICHIESTA')).match(/[a-zA-Z]+/)?.[0]) || 'Kg';
                    const giacenzaMagazzino = parseNumericValue(getVal('GIACENZA_MAGAZZINO'));
                    let qtyToProduce = parseNumericValue(getVal('QTA_DA_PRODURRE'));
                    const needsProduction = !isNaN(qtyRequested) && !isNaN(giacenzaMagazzino) && giacenzaMagazzino < qtyRequested;
                    if (needsProduction && (isNaN(qtyToProduce) || qtyToProduce <= 0)) {
                        qtyToProduce = Math.max(0, qtyRequested - giacenzaMagazzino);
                    }
                    if (qtyToProduce < 0) qtyToProduce = 0;

                    const lottoSCRaw = String(getVal('LOTTO_SC')).trim();
                    const lottoSC = lottoSCRaw.match(/\d+/)?.[0] || '';
                    const produzioneData = parseDateValue(getVal('DATA_PRODUZIONE'));
                    const dataConfezionamento = parseDateValue(getVal('DATA_CONFEZIONAMENTO'));
                    const dataSpedizione = parseDateValue(getVal('DATA_SPEDIZIONE'));

                    const rowData = {
                        codice: String(rawCodice).trim(),
                        prodotto: String(rawProdotto).trim(),
                        cliente: String(getVal('CLIENTE')).trim(),
                        quantitaRichiesta: qtyRequested,
                        quantitaRichiestaUnit: qtyRequestedUnit,
                        giacenzaMagazzino: giacenzaMagazzino === '' ? 0 : giacenzaMagazzino,
                        quantitaDaProdurre: isNaN(qtyToProduce) || qtyToProduce === null ? '' : qtyToProduce,
                        materiePrime: '',
                        macchinari: '',
                        operatore: '',
                        confezionamentoPezzi: packagingDetails.pezzi,
                        confezionamentoKgPerPiece: packagingDetails.kgPerPezzo,
                        confezionamentoUnit: packagingDetails.unit,
                        rawConfezionamentoDettaglio: '',
                        produzioneData: produzioneData,
                        giorniDiProduzione: '',
                        dataConfezionamento: dataConfezionamento,
                        codiceConfezionamento: '',
                        lottoSC: lottoSC,
                        materialeConfezionamento: '',
                        dataSpedizione: dataSpedizione,
                        note: String(getVal('NOTE')).trim()
                    };

                    rowData.macchinari = assignMachine(rowData.codice, rowData.quantitaDaProdurre, rowData.quantitaRichiesta, rowData.giacenzaMagazzino);

                    importedData.push(rowData);
                });
            });

        } catch (error) {
            console.error("Error during Excel file import:", error);
            await showAlert(`Si è verificato un errore durante l'importazione del file Excel: ${error.message}. Controlla la console per maggiori dettagli. Assicurati che il file sia un Excel valido e che le intestazioni delle colonne siano corrette (riga ${headerRowIndex !== -1 ? headerRowIndex + 1 : 'sconosciuta'}).`);
            return;
        }

        if (importedData.length > 0) {
            await compareAndApplyPPChanges(importedData, file.name);
            autoSaveAllData(); // Salva immediatamente i dati nel localStorage
            // Aggiorna il timestamp dell'ultimo import PP e sincronizza la sezione riassuntiva
            try {
                var nowStr;
                if (typeof formatDateTimeForDisplay === 'function') {
                    nowStr = formatDateTimeForDisplay(new Date());
                } else {
                    // Fallback: usa ISO 8601 se la funzione non è disponibile
                    nowStr = new Date().toISOString();
                }
                // Aggiorna chiave uniforme per l'ultimo import PP
                localStorage.setItem('lastImportPP', nowStr);
                if (typeof updateImportTimestamps === 'function') {
                    updateImportTimestamps();
                }
            } catch (errTimestamp) {
                console.warn('Impossibile aggiornare lastImportPP:', errTimestamp);
            }
        } else {
            await showAlert('Nessun dato valido trovato nel file importato o il file è vuoto.');
        }
    }

    async function processOVFile(file) {
        addLogEntry(`--- Inizio importazione Ordini di Vendita (OV) da Excel: ${file.name} ---`);
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false });
            await transformAndImportOVData(json, file.name);
        } catch (error) {
            console.error("Errore durante l'importazione del file Excel OV:", error);
            addLogEntry(`Importazione OV fallita: ${error.message}`);
            await showAlert(`Errore durante l'importazione del file Excel OV: ${error.message}.`);
        }
    }

    async function processOVFileFromCSV(file) {
        addLogEntry(`--- Inizio importazione Ordini di Vendita (OV) da CSV: ${file.name} ---`);
        try {
            const text = await file.text();
            const rows = text.split('\n').map(row =>
                row.trim().split(';').map(cell => cell.trim())
            );

            while (rows.length > 0 && (rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '')) {
                rows.pop();
            }

            if (rows.length === 0) {
                await showAlert('Il file CSV è vuoto o non contiene dati validi.');
                return;
            }

            await transformAndImportOVData(rows, file.name);

        } catch (error) {
            console.error("Errore durante l'importazione del file CSV OV:", error);
            addLogEntry(`Importazione CSV OV fallita: ${e.message}`);
            await showAlert(`Errore durante l'importazione del file CSV OV: ${e.message}.`);
        }
    }
    // Cerca la funzione transformAndImportOVData e sostituiscila con questa versione corretta:

    async function transformAndImportOVData(data, fileName) {
        const importedData = [];

        // NUOVA FUNZIONE CORRETTA per calcolare l'indice delle colonne Excel (es. 'BH')
        const excelColToIndex = (colName) => {
            let index = 0;
            for (let i = 0; i < colName.length; i++) {
                index *= 26;
                index += colName.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
            }
            return index - 1;
        };

        // FIX 2: Utilizzo della nuova funzione corretta
        const colIndices = {
            'ov': excelColToIndex('BH'),
            'codice': excelColToIndex('I'),
            'descrizione': excelColToIndex('C'),
            'quantita': excelColToIndex('M'),
            'unitaMisura': excelColToIndex('N'),
            'dataConsegna': excelColToIndex('L'),
            'dataRichiestaCliente': excelColToIndex('J'),
            'dataConferma': excelColToIndex('K'),
        };

        const dataRows = data.filter(row => row && row.length > 10);

        dataRows.forEach(async (row) => {
            const ovValue = String(row[colIndices['ov']] || '').trim();

            // FIX 1: Salta la riga se il valore nella colonna OV non è un numero (scarta le intestazioni)
            if (!ovValue || isNaN(parseInt(ovValue, 10))) {
                return;
            }

            const rowData = {
                ov: ovValue,
                codice: String(row[colIndices['codice']] || '').trim(),
                descrizione: String(row[colIndices['descrizione']] || '').trim(),
                quantitaOrdine: parseNumericValue(row[colIndices['quantita']]),
                unitaMisura: String(row[colIndices['unitaMisura']] || '').trim(),
                dataConsegna: parseDateValue(row[colIndices['dataConsegna']]),
                dataRichiestaCliente: parseDateValue(row[colIndices['dataRichiestaCliente']]),
                dataConferma: parseDateValue(row[colIndices['dataConferma']]),
                note: ''
            };

            if (rowData.ov && rowData.codice && rowData.descrizione) {
                importedData.push(rowData);
            }
        });

        if (importedData.length > 0) {
            let addedCount = 0;
            for (const newRowData of importedData) {
                const canAdd = await handlePossibleDuplicateOV(newRowData);
                if (canAdd) {
                    salesOrderTableBody.appendChild(createSalesOrderRow(newRowData));
                    addedCount++;
                }
            }
            await saveDataToServer();
            await showAlert(`Importazione completata. Aggiunte ${addedCount} nuove righe di ordini di vendita dal file "${fileName}".`);
            addLogEntry(`Importazione OV completata: aggiunte ${addedCount} righe da "${fileName}".`);
            runFullCheck();
        } else {
            await showAlert('Nessun dato valido trovato nel file OV importato o il formato non corrisponde.');
            addLogEntry(`Importazione OV fallita: nessun dato valido trovato in "${fileName}".`);
        }
    }


async function processOpiFile(file) {
    const data = new Uint8Array(await file.arrayBuffer());
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });

    // Trova la riga header dove c’è “OP” come primo valore
    let headerRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
        if ((rows[i][0] || "").toString().toUpperCase().includes('OP')) {
            headerRowIndex = i;
            break;
        }
    }
    const dataRows = headerRowIndex >= 0 ? rows.slice(headerRowIndex + 1) : rows.slice(1);

    // Mappatura colonne aggiornata (modifica se necessario!)
    const colMap = {
        dataProd: 2,   // C
        op: 1,         // B
        ov: 13,        // N
        codice: 3,     // D
        articolo: 4,   // E
        cliente: 12,   // M
        lotto: 8,      // I
        quantita: 6,   // G
        um: 5,         // F
        operatore: 92, // O
        scadenza: 25   // Z
    };

    const opiData = [];
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        if (!row || (!row[colMap.op] && !row[colMap.codice])) continue;
        opiData.push({
            op: row[colMap.op] || "",
            dataProd: row[colMap.dataProd] || "",
            codice: row[colMap.codice] || "",
            articolo: row[colMap.articolo] || "",
            um: row[colMap.um] || "",
            operatore: row[colMap.operatore] || "",
            quantita: row[colMap.quantita] || "",
            lotto: row[colMap.lotto] || "",
            cliente: row[colMap.cliente] || "",
            ov: row[colMap.ov] || "",
            scadenza: row[colMap.scadenza] || ""
        });
    }

    populateOpiTable(opiData);
    await saveOpiDataToLocalAndServer(opiData);
    await showAlert(`Importazione OPI completata. Righe importate: ${opiData.length}.`);
}


  async function processOSFile(file, targetTableBody, rowCreationFunction) {
    // CORREZIONE 1: Identifica correttamente se l'importazione è per 'Arrivi' o 'OS'
    const logPrefix = targetTableBody.parentElement.id === 'arrivalScheduleTable' ? 'Arrivi' : 'OS';
    addLogEntry(`--- Inizio importazione ${logPrefix} da: ${file.name} ---`);

    try {
        const data = new Uint8Array(await file.arrayBuffer());
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });

        const excelColToIndex = (col) => {
            let index = 0;
            for (let i = 0; i < col.length; i++) {
                index = index * 26 + col.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
            }
            return index - 1;
        };

        const colMap = {
            ov: excelColToIndex('L'),
            codiceArticolo: excelColToIndex('B'),
            descrizioneArticolo: excelColToIndex('C'),
            quantita: excelColToIndex('E'),
            um: excelColToIndex('D'),
            dataConsegna: excelColToIndex('J'),
            dataConferma: excelColToIndex('V'),
            ragioneSociale: excelColToIndex('N'),
            riferimentoCliente: excelColToIndex('CW'),
            indirizzo: excelColToIndex('DA'),
            cap: excelColToIndex('DB'),
            citta: excelColToIndex('DC'),
            provincia: excelColToIndex('DD'),
            telefono: excelColToIndex('DE'),
            // La colonna Q del file (lettera "Q") contiene note di servizio/interne. Questa
            // informazione viene importata ma non mostrata nella tabella, sarà visibile
            // solo nel tooltip quando si passa con il mouse sulla riga.
            noteServizio: excelColToIndex('Q'),
            // La colonna CU identifica lo stato di conferma arrivo (es. "O", "C", ecc.).
            // Non viene mostrata in tabella, ma serve per decidere se importare la riga
            // nella tabella degli arrivi. Solo le righe con valore "O" o "C" vengono
            // importate; le altre vengono ignorate (regola merce non arrivata/quarantena).
            cu: excelColToIndex('CU')
        };

        const importedData = [];
        jsonData.forEach((row, index) => {
            if (index === 0 || row.every(cell => cell === "")) return;
            const codiceArticolo = String(row[colMap.codiceArticolo] || '').trim();
            const layoutInfo = layoutData[codiceArticolo] || { layout: '', family: 'Senza Famiglia' };
            // Determina il valore della colonna CU (stato di conferma arrivo). Non viene
            // mostrato in tabella, ma è utilizzato per filtrare le righe importate.
            const cuValRaw = row[colMap.cu];
            const cuVal = cuValRaw !== undefined && cuValRaw !== null ? String(cuValRaw).trim().toUpperCase() : '';
            // Crea l'oggetto riga con i campi noti. Il valore CU non viene incluso nella
            // struttura, ma può essere usato per filtrare (solo per gli arrivi).
            const rowData = {
                ov: String(row[colMap.ov] || '').trim(),
                codiceArticolo: codiceArticolo,
                descrizioneArticolo: String(row[colMap.descrizioneArticolo] || '').trim(),
                quantita: parseNumericValue(row[colMap.quantita]),
                um: String(row[colMap.um] || '').trim(),
                dataConsegna: parseDateValue(row[colMap.dataConsegna]),
                dataConferma: parseDateValue(row[colMap.dataConferma]),
                ragioneSociale: String(row[colMap.ragioneSociale] || '').trim(),
                riferimentoCliente: String(row[colMap.riferimentoCliente] || '').trim(),
                indirizzo: String(row[colMap.indirizzo] || '').trim(),
                cap: String(row[colMap.cap] || '').trim(),
                citta: String(row[colMap.citta] || '').trim(),
                provincia: String(row[colMap.provincia] || '').trim(),
                telefono: String(row[colMap.telefono] || '').trim(),
                layout: layoutInfo.layout,
                family: layoutInfo.family,
                noteServizio: String(row[colMap.noteServizio] || '').trim()
            };
            // Aggiungi la riga ai dati importati solo se contiene informazioni utili.
            // Se stiamo importando gli Arrivi, filtra in base al valore della colonna CU:
            // solo se la CU è "O" o "C" si considera la riga (merce confermata o controllata).
            if (rowData.codiceArticolo || rowData.ov) {
                if (logPrefix === 'Arrivi') {
                    if (cuVal === 'O' || cuVal === 'C') {
                        importedData.push(rowData);
                    }
                } else {
                    importedData.push(rowData);
                }
            }
        });

        if (logPrefix === 'Arrivi') {
            const today = new Date();
            today.setHours(0, 0, 0, 0);

            const currentArrivals = [];
            const overdueArrivals = [];

            importedData.forEach(item => {
                // Filtra gli articoli di servizio/varie o con ragione sociale di laboratorio.
                // Se la descrizione o la ragione sociale contengono questi indicatori,
                // non includere l'articolo nella lista "Merce non Arrivata".  Si usa
                // un test case-insensitive con corrispondenza parziale.
                const desc = String(item.descrizioneArticolo || '').toLowerCase();
                const ragSoc = String(item.ragioneSociale || '').toLowerCase();
                const unwantedDescRegex = /servizi\s*e\s*varie|profilo\s+micro\w*\s+esteso|gestione\s+campione/;
                const unwantedRagSocRegex = /chelab(?:\s*srl)?|lab4life(?:\s*srl)?/;
                if (unwantedDescRegex.test(desc) || unwantedRagSocRegex.test(ragSoc)) {
                    // Ignora completamente questi articoli di servizio, non vanno né tra gli arrivi
                    // né tra la merce non arrivata.
                    return;
                }

                const deliveryDateParts = item.dataConsegna.split('/');
                if (deliveryDateParts.length !== 3) {
                    currentArrivals.push(item);
                    return;
                }
                const deliveryDate = new Date(parseInt(deliveryDateParts[2]), parseInt(deliveryDateParts[1]) - 1, parseInt(deliveryDateParts[0]));

                if (deliveryDate < today) {
                    overdueArrivals.push(item);
                } else {
                    currentArrivals.push(item);
                }
            });
            
            // CORREZIONE 2: Popola entrambe le tabelle con i dati smistati
            arrivalScheduleTableBody.innerHTML = '';
            currentArrivals.forEach(newRowData => {
                arrivalScheduleTableBody.appendChild(createArrivalScheduleRow(newRowData));
            });

            populateOverdueTable(overdueArrivals);

            addLogEntry(`Importazione arrivi completata. Righe correnti: ${currentArrivals.length}, Righe scadute: ${overdueArrivals.length}.`);
            await showAlert(`Importazione completata. Trovati ${currentArrivals.length} arrivi programmati e ${overdueArrivals.length} articoli non ancora arrivati.`);

        } else { // Logica per le Spedizioni (OS)
            targetTableBody.innerHTML = '';
            importedData.forEach(newRowData => {
                targetTableBody.appendChild(rowCreationFunction(newRowData));
            });
            await showAlert(`Importazione ${logPrefix} completata. Tabella aggiornata con ${importedData.length} righe.`);
        }

        // Aggiorna il gantt delle spedizioni/arrivi dopo l'importazione
        updateWarehouseGanttChart();
        // Salva i dati sul server e localmente
        await saveDataToServer();
        autoSaveAllData();
        // Esegui immediatamente il controllo ADR dopo l'importazione. In
        // questo modo eventuali spedizioni ADR appena importate
        // verranno segnalate senza attendere il refresh periodico.  Il
        // controllo viene effettuato solo se la funzione è definita.
        if (typeof checkAndNotifyADR === 'function') {
            try {
                checkAndNotifyADR();
            } catch (e) {
                console.warn('Errore nel controllo ADR dopo importazione OS/Arrivi:', e);
            }
        }

    } catch (error) {
        console.error(`Errore durante l'importazione del file ${logPrefix}:`, error);
        addLogEntry(`Importazione ${logPrefix} fallita: ${error.message}`);
        await showAlert(`Errore durante l'importazione del file ${logPrefix}: ${error.message}.`);
    }
}

function populateOpiTable(opiData) {
    opiTableBody.innerHTML = '';
    opiData.forEach(data => {
        opiTableBody.appendChild(createOpiRow(data));
    });
    // Dopo aver popolato la tabella OPI, aggiorna i riferimenti nella tabella giornaliera
    if (typeof updateDailyOpeOv === 'function') {
        try {
            updateDailyOpeOv();
        } catch (err) {
            console.warn('Errore durante l\'aggiornamento automatico delle colonne OPE/OV dalla tabella OPI:', err);
        }
    }
}

function createOpiRow(data) {
    // Triade identificativa
    const triadeKey = `${data.op || ""}_${data.ov || ""}_${data.lotto || ""}`;
    const row = document.createElement('tr');
    row.innerHTML = `
        <td>${data.dataProd}</td>
        <td>${data.op}</td>
        <td>${data.ov}</td>
        <td>${data.codice}</td>
        <td>${data.articolo}</td>
        <td>${data.cliente}</td>
        <td>${data.lotto}</td>
        <td>${data.quantita}</td>
        <td>${data.um}</td>
        <td>${data.operatore}</td>
        <td>${data.scadenza}</td>
        <td>
            <span class="log-lens" title="Visualizza log movimenti" style="cursor:pointer;font-size:1.3em;" data-triade="${triadeKey}">
                🔍
            </span>
        </td>
    `;
    // Lente: quando clicchi, mostra il log della triade
    row.querySelector('.log-lens').addEventListener('click', function(e) {
        showOpiMovementLogModal(triadeKey);
    });
    return row;
}

// Struttura log movimenti OPI
let opiMovementLog = JSON.parse(localStorage.getItem('opi_movement_log') || '[]');

// Salva log su localStorage
function saveOpiMovementLog() {
    localStorage.setItem('opi_movement_log', JSON.stringify(opiMovementLog));
}

// Logga una variazione OPI
function logOpiMovement(triadeKey, changeList) {
    const now = new Date().toISOString();
    opiMovementLog.push({
        triade: triadeKey,
        timestamp: now,
        changes: changeList
    });
    saveOpiMovementLog();
}

// Mostra il log movimenti OPI (ultimi 7 giorni)
function showOpiMovementLogModal(triadeKey) {
    const now = new Date();
    const oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const logEntries = opiMovementLog.filter(entry =>
        entry.triade === triadeKey &&
        new Date(entry.timestamp) >= oneWeekAgo
    );
    let html = '';
    if (logEntries.length === 0) {
        html = '<p>Nessun movimento registrato negli ultimi 7 giorni.</p>';
    } else {
        html = logEntries.map(entry => {
            return `<div style="margin-bottom:8px;">
                <strong>${new Date(entry.timestamp).toLocaleString('it-IT')}</strong><br>
                ${entry.changes.map(c =>
                    `Campo <b>${c.field}</b>: <span style="color:red;">${c.oldValue}</span> → <span style="color:green;">${c.newValue}</span>`
                ).join('<br>')}
            </div>`;
        }).join('');
    }
    showCustomModal('Log movimenti OPI', html, [{text: 'Chiudi', class: 'alert', value: true}]);
}

// Esempio di tracciamento: chiamare questa funzione ogni volta che modifichi una riga OPI
// logOpiMovement(triadeKey, [{field:'quantita', oldValue:'10', newValue:'12'}]);

function populateOpiTable(opiData) {
    opiTableBody.innerHTML = '';
    opiData.forEach(data => {
        opiTableBody.appendChild(createOpiRow(data));
    });
}

function createOpiRow(data) {
    const row = document.createElement('tr');
    row.innerHTML = `
        <td>${data.dataProd}</td>
        <td>${data.op}</td>
        <td>${data.ov}</td>
        <td>${data.codice}</td>
        <td>${data.articolo}</td>
        <td>${data.cliente}</td>
        <td>${data.lotto}</td>
        <td>${data.quantita}</td>
        <td>${data.um}</td>
        <td>${data.operatore}</td>
        <td>${data.scadenza}</td>
        <td></td>
    `;
    return row;
}

    // Sostituiamo l'utilizzo di fileInput per l'importazione del programma di produzione
    // con un input dinamico.  Questo evita problemi di puntatori e permette di
    // selezionare il file in maniera affidabile anche quando il fileInput
    // principale è disabilitato o nascosto.  importMode è mantenuto per
    // compatibilità con la logica esistente ma non viene più utilizzato per il click.
    importPPBtn.addEventListener('click', () => {
        importMode = 'PP';
        const dynInput = document.createElement('input');
        dynInput.type = 'file';
        dynInput.accept = '.xls,.xlsx';
        dynInput.style.display = 'none';
        document.body.appendChild(dynInput);
        dynInput.addEventListener('change', async (e) => {
            const file = e.target.files && e.target.files[0];
            if (!file) {
                dynInput.remove();
                return;
            }
            const ext = file.name.split('.').pop().toLowerCase();
            const isExcel = (ext === 'xls' || ext === 'xlsx');
            if (isExcel) {
                try {
                    await processPPFile(file);
                } catch (err) {
                    console.error('Errore durante l\'importazione del file PP:', err);
                    if (typeof showAlert === 'function') {
                        await showAlert(`Errore durante l'importazione del file PP: ${err.message}`);
                    }
                }
            } else {
                if (typeof showAlert === 'function') {
                    await showAlert('Formato file non supportato per PP. Seleziona un file Excel (.xls, .xlsx).');
                }
            }
            dynInput.remove();
        });
        dynInput.click();
    });

    importOVBtn.addEventListener('click', () => {
        importMode = 'OV';
        fileInput.click();
    });

   // Gestione del click per il bottone di importazione OS spostata nella sezione
   // "ATTIVAZIONE EVENTI" più in basso per evitare duplicazioni.  Qui non
   // assegniamo alcun handler.

importArrivalsBtn.addEventListener('click', () => {
        // Gestisce l'importazione degli arrivi con un input file dedicato.  La
        // logica è separata dalla variabile importMode per evitare conflitti con
        // altri gestori.  Viene creata una lista dinamica per selezionare il
        // file, che dopo l'uso viene rimossa.
        const dynInput = document.createElement('input');
        dynInput.type = 'file';
        dynInput.accept = '.xls,.xlsx,.csv';
        dynInput.style.display = 'none';
        document.body.appendChild(dynInput);
        dynInput.addEventListener('change', async (evt) => {
            const file = evt.target.files[0];
            if (!file) {
                dynInput.remove();
                return;
            }
            const ext = file.name.split('.').pop().toLowerCase();
            const isExcel = ext === 'xls' || ext === 'xlsx';
            const isCsv = ext === 'csv';
            if (isExcel || isCsv) {
                await processOSFile(file, arrivalScheduleTableBody, createArrivalScheduleRow);
                // Aggiorna timestamp dell'ultimo import Arrivi
                const ts = formatDateTimeForDisplay(new Date());
                try {
                    // Usa una chiave uniforme senza underscore
                    localStorage.setItem('lastImportArrivals', ts);
                } catch (e) {}
                if (typeof updateImportTimestamps === 'function') updateImportTimestamps();
            } else {
                await showAlert('Formato file non supportato per Arrivi. Seleziona un file Excel o CSV.');
            }
            dynInput.remove();
        });
        dynInput.click();
    });

    importOpiBtn.addEventListener('click', () => {
    importMode = 'OPI';
    fileInput.click();
});

    // Gestione del click per importare il file DeviceRef.  Imposta l'importMode
    // a 'deviceRef' e richiama il fileInput per selezionare il file Excel/CSV.
    if (importDeviceRefBtn) {
        importDeviceRefBtn.addEventListener('click', () => {
            importMode = 'deviceRef';
            fileInput.click();
        });
    }

    importReferenzeBtn.addEventListener('click', () => {
        importMode = 'referenze';
        fileInput.click(); // CORRETTO
    });

    importPianoAnaliticoBtn.addEventListener('click', () => {
        importMode = 'pianoAnalitico';
        fileInput.click(); // CORRETTO
    });

    fileInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const fileExtension = file.name.split('.').pop().toLowerCase();
        const isExcel = fileExtension === 'xls' || fileExtension === 'xlsx';
        const isCsv = fileExtension === 'csv';

        switch (importMode) {
            case 'PP':
                if (isExcel) {
                    await processPPFile(file);
                } else {
                    await showAlert('Formato file non supportato per PP. Seleziona un file Excel (.xls, .xlsx).');
                }
                break;
            case 'OV':
                if (isExcel) {
                    await processOVFile(file);
                } else if (isCsv) {
                    await processOVFileFromCSV(file);
                } else {
                    await showAlert('Formato file non supportato per OV. Seleziona un file Excel (.xls, .xlsx) o CSV (.csv).');
                }
                break;
             case 'OPI':
                 if (isExcel || isCsv) {
                   await processOpiFile(file);
                } else {
                   await showAlert('Formato file non supportato per OPI. Seleziona un file Excel o CSV.');
                }
                break;
            case 'OS':
                if (isExcel || isCsv) {
                    // La funzione ora sa quale tabella usare
                    await processOSFile(file, shippingScheduleTableBody, createShippingScheduleRow);
                    // Rende ordinabile la tabella spedizioni anche dopo un import
                    if (typeof makeTableSortable === 'function') {
                        makeTableSortable(document.getElementById('shippingScheduleTable'));
                    }
                } else {
                    await showAlert('Formato file non supportato per OS. Seleziona un file Excel o CSV.');
                }
                break;
           case 'Arrivals':
                if (isExcel || isCsv) {
                    // Riusiamo la stessa funzione di importazione passando la tabella degli arrivi
                    await processOSFile(file, arrivalScheduleTableBody, createArrivalScheduleRow);
                    // Rendi ordinabile la tabella arrivi anche dopo un import
                    if (typeof makeTableSortable === 'function') {
                        makeTableSortable(document.getElementById('arrivalScheduleTable'));
                    }
                } else {
                    await showAlert('Formato file non supportato per Arrivi. Seleziona un file Excel o CSV.');
                }
                break;
           case 'Layout':
                if (isExcel || isCsv) {
                    await processLayoutFile(file);
                } else {
                    await showAlert('Formato file non supportato per Layout. Seleziona un file Excel o CSV.');
                }
                break;
            case 'referenze': // NUOVA LOGICA AGGIUNTA
                if (isExcel || isCsv) {
                    await processReferenzeFile(file);
                } else {
                    await showAlert('Formato file non supportato. Seleziona un file Excel o CSV.');
                }
                break;
            case 'pianoAnalitico': // NUOVA LOGICA AGGIUNTA
                if (isExcel || isCsv) {
                    await processPianoAnaliticoFile(file);
                } else {
                    await showAlert('Formato file non supportato. Seleziona un file Excel o CSV.');
                }
                break;
            case 'deviceRef': // Gestione dell'import dei riferimenti dispositivi/medicali
                if (isExcel || isCsv) {
                    await processDeviceRefFile(file);
                } else {
                    await showAlert('Formato file non supportato. Seleziona un file Excel o CSV.');
                }
                break;
        }

        // Dopo aver processato il file, registra la data/ora di import per il tipo di file corrente.
        if (importMode) {
            const nowTs = formatDateTimeForDisplay(new Date());
            try {
                // Utilizza una chiave uniforme senza underscore: es. lastImportOV, lastImportOS
                // Costruisce la chiave uniformemente in CamelCase: es.
                // 'OS' -> 'lastImportOS', 'deviceRef' -> 'lastImportDeviceRef'
                const normalized = importMode.charAt(0).toUpperCase() + importMode.slice(1);
                const key = 'lastImport' + normalized;
                localStorage.setItem(key, nowTs);
            } catch (e) {
                // In caso di impossibilità a scrivere sul localStorage (es. modalità privata), ignora
            }
            updateImportTimestamps();
        }

        fileInput.value = '';
        importMode = null;
    });
    // Gestore per il caricamento del file Referenze
    referenzeInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            referenzeData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            saveStaticData('referenzeData', referenzeData, file.name);

            // Aggiorna l'etichetta con il nome del file ma non mostra più il flag visivo.
            referenzeFileStatusSpan.textContent = `File: ${file.name}`;
            referenzeFileStatusSpan.style.display = 'none';
            await showAlert('Nuovo file Referenze importato con successo!');
            await saveDataToServer();
            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
            renderAnalysisTableHeaders();
            updateAnalisiTable();

        } catch (e) {
            await showAlert(`Errore durante l'importazione del file Referenze: ${e.message}`);
        }
    });

    // Gestore per il caricamento del file Piano Analitico
    pianoAnaliticoInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            pianoAnaliticoData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            saveStaticData('pianoAnaliticoData', pianoAnaliticoData, file.name);

            // Aggiorna l'etichetta con il nome del file ma non mostra più il flag visivo
            pianoAnaliticoFileStatusSpan.textContent = `File: ${file.name}`;
            pianoAnaliticoFileStatusSpan.style.display = 'none';
            await showAlert('Nuovo Piano Analitico importato con successo!');
            await saveDataToServer();
            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
            renderAnalysisTableHeaders();
            updateAnalisiTable();

        } catch (e) {
            await showAlert(`Errore durante l'importazione del file Piano Analitico: ${e.message}`);
        }
    });

    
// =================================================================================
// BLOCCO DI CORREZIONE DEFINITIVA PER I FLAG DI ANALISI - INCOLLARE ALLA FINE DELLO SCRIPT
// =================================================================================

/**
 * Funzione CORRETTA che processa il file Referenze.
 * Mostra solo il flag visivo al successo.
 */
async function processReferenzeFile(file) {
    try {
        const data = new Uint8Array(await file.arrayBuffer());
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        referenzeData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
        saveStaticData('referenzeData', referenzeData, file.name);

        const flag = document.getElementById('referenzeFileStatus');
        if (flag) {
            flag.style.display = 'flex';
        }

        await showAlert('Nuovo file Referenze importato con successo!');
        loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
        renderAnalysisTableHeaders();
        updateAnalisiTable();
    } catch (e) {
        await showAlert(`Errore durante l'importazione del file Referenze: ${e.message}`);
    }
}

/**
 * Funzione CORRETTA che processa il file Piano Analitico.
 * Mostra solo il flag visivo al successo.
 */
async function processPianoAnaliticoFile(file) {
    try {
        const data = new Uint8Array(await file.arrayBuffer());
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        pianoAnaliticoData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
        saveStaticData('pianoAnaliticoData', pianoAnaliticoData, file.name);

        const flag = document.getElementById('pianoAnaliticoFileStatus');
        if (flag) {
            flag.style.display = 'flex';
        }

        await showAlert('Nuovo Piano Analitico importato con successo!');
        loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
        renderAnalysisTableHeaders();
        updateAnalisiTable();
    } catch (e) {
        await showAlert(`Errore durante l'importazione del file Piano Analitico: ${e.message}`);
    }
}

/**
 * Processa un file Excel o CSV contenente le informazioni di riferimento per i
 * dispositivi/medicali (DeviceRef).  Ogni riga della tabella contiene un
 * codice articolo nella colonna A, il cliente in colonna B, una colonna C
 * ignorata, quindi varie informazioni tecniche nelle colonne successive:
 * colonna D indica se sono presenti aghi, la E la tipologia di aghi, la F
 * il numero di aghi per valva, la G il numero di siringhe per scatola,
 * la H il volume della siringa in ml, la I il numero di siringhe per scatola
 * (secondo formato), la J il peso di ogni singola scatola e la K il peso
 * dello scatolone complessivo.  I dati vengono memorizzati in localStorage
 * con chiave "deviceRefData" e inviati al server tramite saveDataToServer().
 * Viene mostrato un messaggio di conferma al termine.
 *
 * @param {File} file - Il file Excel o CSV da elaborare
 */
async function processDeviceRefFile(file) {
    try {
        const data = new Uint8Array(await file.arrayBuffer());
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
        const deviceRefs = [];
        // Itera sulle righe, ignorando la prima riga di intestazione se presente
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            // Colonne: A=0 (codice), B=1 (cliente), C=2 (ignored), D=3 (aghi presenti),
            // E=4 (tipologia aghi), F=5 (aghi per valva), G=6 (siringhe per scatola),
            // H=7 (volume ml), I=8 (siringhe per scatola 2), J=9 (peso scatola),
            // K=10 (peso scatolone)
            const codice = row[0] ? row[0].toString().trim() : '';
            if (!codice) continue;
            const cliente = row[1] ? row[1].toString().trim() : '';
            const aghiPresenti = row[3] ? row[3].toString().trim() : '';
            const tipologiaAghi = row[4] ? row[4].toString().trim() : '';
            const aghiPerValva = row[5] ? row[5].toString().trim() : '';
            const siringhePerScatola = row[6] ? row[6].toString().trim() : '';
            const volumeMl = row[7] ? row[7].toString().trim() : '';
            const siringhePerScatola2 = row[8] ? row[8].toString().trim() : '';
            const pesoScatola = row[9] ? row[9].toString().trim() : '';
            const pesoScatolone = row[10] ? row[10].toString().trim() : '';
            deviceRefs.push({
                codice: codice.toUpperCase(),
                cliente,
                aghiPresenti,
                tipologiaAghi,
                aghiPerValva,
                siringhePerScatola,
                volumeMl,
                siringhePerScatola2,
                pesoScatola,
                pesoScatolone
            });
        }
        // Salva su localStorage
        localStorage.setItem('deviceRefData', JSON.stringify(deviceRefs));
        // Salva anche sul server inclusa questa informazione
        await saveDataToServer();
        // Mostra feedback all'utente
        await showAlert('Nuovo file DeviceRef importato con successo!');
        // Aggiorna la data/ora dell'ultimo import DeviceRef sia nel localStorage
        // che nell'etichetta vicino al bottone
        const nowStr = formatDateTimeForDisplay(new Date());
        try {
            localStorage.setItem('lastImportDeviceRef', nowStr);
        } catch (e) {}
        const statusSpan = document.getElementById('lastImportDeviceRef');
        if (statusSpan) {
            statusSpan.textContent = ` (Ultimo import: ${nowStr})`;
        }
        if (typeof updateImportTimestamps === 'function') {
            updateImportTimestamps();
        }
    } catch (e) {
        console.error('Errore durante l\'importazione del file DeviceRef:', e);
        await showAlert(`Errore durante l'importazione del file DeviceRef: ${e.message}`);
    }
}

/**
 * Funzione CORRETTA che carica i dati all'avvio della pagina.
 * Mostra i flag SOLO per i file effettivamente salvati in memoria.
 */
function loadStaticData() {
    try {
        const savedReferenze = localStorage.getItem('referenzeData');
        const savedPianoAnalitico = localStorage.getItem('pianoAnaliticoData');

        // Gestisce il flag per Referenze in modo indipendente
        const referenzeFlag = document.getElementById('referenzeFileStatus');
        if (savedReferenze && referenzeFlag) {
            referenzeData = JSON.parse(savedReferenze);
            referenzeFlag.style.display = 'flex';
        }

        // Gestisce il flag per Piano Analitico in modo indipendente
        const pianoAnaliticoFlag = document.getElementById('pianoAnaliticoFileStatus');
        if (savedPianoAnalitico && pianoAnaliticoFlag) {
            pianoAnaliticoData = JSON.parse(savedPianoAnalitico);
            pianoAnaliticoFlag.style.display = 'flex';
        }

        if (savedReferenze && savedPianoAnalitico) {
            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
        }
    } catch (e) {
        console.error("Errore nel caricamento dei file di analisi da localStorage:", e);
    }
}
// =================================================================================
// FINE BLOCCO DI CORREZIONE
// =================================================================================

    fileInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const fileExtension = file.name.split('.').pop().toLowerCase();
        const isExcel = fileExtension === 'xls' || fileExtension === 'xlsx';
        const isCsv = fileExtension === 'csv';

        switch (importMode) {
            case 'PP':
                if (isExcel) {
                    await processPPFile(file);
                } else {
                    await showAlert('Formato file non supportato per PP. Seleziona un file Excel (.xls, .xlsx).');
                }
                break;
            case 'OV':
                if (isExcel) {
                    await processOVFile(file);
                } else if (isCsv) {
                    await processOVFileFromCSV(file);
                } else {
                    await showAlert('Formato file non supportato per OV. Seleziona un file Excel (.xls, .xlsx) o CSV (.csv).');
                }
                break;
            case 'referenze':
                if (isExcel || isCsv) {
                    await processReferenzeFile(file);
                } else {
                    await showAlert('Formato file non supportato. Seleziona un file Excel o CSV.');
                }
                break;
            case 'pianoAnalitico':
                if (isExcel || isCsv) {
                    await processPianoAnaliticoFile(file);
                } else {
                    await showAlert('Formato file non supportato. Seleziona un file Excel o CSV.');
                }
                break;
        }

        fileInput.value = '';
        importMode = null;
    });
    async function processReferenzeFile(file) {
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            referenzeData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            saveStaticData('referenzeData', referenzeData, file.name);

            // Aggiorna etichetta del file importato (nascosta) e avvisa l'utente
            if (referenzeFileStatusSpan) {
                referenzeFileStatusSpan.textContent = `File: ${file.name}`;
                referenzeFileStatusSpan.style.display = 'none';
            }
            await showAlert('Nuovo file Referenze importato con successo!');

            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
            renderAnalysisTableHeaders();
            updateAnalisiTable();

            // Registra la data/ora dell'ultimo import Referenze su una chiave uniforme
            if (typeof formatDateTimeForDisplay === 'function') {
                const nowStr = formatDateTimeForDisplay(new Date());
                localStorage.setItem('lastImportReferenze', nowStr);
            } else {
                localStorage.setItem('lastImportReferenze', Date.now().toString());
            }
            if (typeof updateImportTimestamps === 'function') {
                updateImportTimestamps();
            }
        } catch (e) {
            await showAlert(`Errore durante l'importazione del file Referenze: ${e.message}`);
        }
    }

    async function processPianoAnaliticoFile(file) {
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            pianoAnaliticoData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            saveStaticData('pianoAnaliticoData', pianoAnaliticoData, file.name);

            // Aggiorna etichetta del file importato (nascosta) e avvisa l'utente
            if (pianoAnaliticoFileStatusSpan) {
                pianoAnaliticoFileStatusSpan.textContent = `File: ${file.name}`;
                pianoAnaliticoFileStatusSpan.style.display = 'none';
            }
            await showAlert('Nuovo Piano Analitico importato con successo!');

            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
            renderAnalysisTableHeaders();
            updateAnalisiTable();

            // Registra la data/ora dell'ultimo import Piano Analitico su una chiave uniforme
            if (typeof formatDateTimeForDisplay === 'function') {
                const nowStr = formatDateTimeForDisplay(new Date());
                localStorage.setItem('lastImportPianoAnalitico', nowStr);
            } else {
                localStorage.setItem('lastImportPianoAnalitico', Date.now().toString());
            }
            if (typeof updateImportTimestamps === 'function') {
                updateImportTimestamps();
            }
        } catch (e) {
            await showAlert(`Errore durante l'importazione del file Piano Analitico: ${e.message}`);
        }
    }

    // ===================================================================
    // ==> FUNZIONI PER L'IMPORTAZIONE DELLA PRODUZIONE MEDICAL DEVICE <==
    // ===================================================================
    /**
     * Importa un file Excel o CSV contenente i dati di produzione dei dispositivi medici.
     * Solo i codici presenti nelle referenze dei dispositivi (deviceRefData) vengono importati.
     * La prima colonna viene interpretata come codice, la colonna D come data
     * (in formato Excel o testo), la colonna M come quantità, la colonna N come
     * unità di misura e la colonna I viene utilizzata per identificare eventuali
     * righe di scarto (contiene la parola "scarto"/"scarti").  Le quantità
     * negative vengono considerate scarti.  I valori di scarto vengono
     * aggregati alla riga principale (quantità positiva) con stessa data e codice.
     * @param {File} file Il file Excel/CSV da importare
     */
    async function processMedicalProductionFile(file) {
        addLogEntry(`--- Inizio importazione Produzione Medical Device da ${file.name} ---`);
        try {
            let rows;
            const fileName = file.name.toLowerCase();
            if (fileName.endsWith('.csv')) {
                const text = await file.text();
                rows = text.split(/\r?\n/).map(row => row.split(';').map(cell => cell.trim()));
            } else {
                const data = new Uint8Array(await file.arrayBuffer());
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false });
            }
            if (!rows || rows.length === 0) {
                await showAlert('Il file importato è vuoto o non contiene dati validi.');
                return;
            }
            await transformAndImportMedicalProductionData(rows, file.name);
            // Aggiorna timestamp ultimo import su chiave uniforme
            const nowStr = formatDateTimeForDisplay(new Date());
            localStorage.setItem('lastImportMedicalProduction', nowStr);
            updateImportTimestamps();
            // Salva su server i dati appena importati
            await saveDataToServer();
            addLogEntry(`--- Fine importazione Produzione Medical Device (${file.name}) ---`);
        } catch (err) {
            console.error('Errore durante l\'importazione Produzione Medical Device:', err);
            addLogEntry(`Importazione Produzione MD fallita: ${err.message}`);
            await showAlert(`Errore durante l'importazione del file Produzione Medical Device: ${err.message}`);
        }
    }

    /**
     * Converte i dati grezzi del foglio in una struttura per la tabella di
     * produzione medicale. Aggrega quantità e scarti per codice e data.
     * @param {Array[]} data Matrice di righe (array di celle) provenienti da Excel/CSV
     * @param {string} fileName Nome del file (per eventuali log)
     */
    async function transformAndImportMedicalProductionData(data, fileName) {
        const deviceRefs = typeof getDeviceRefData === 'function' ? getDeviceRefData() : JSON.parse(localStorage.getItem('deviceRefData') || '[]');
        // Crea un set di codici validi per i dispositivi medici
        const validCodes = new Set(deviceRefs.map(ref => ref.codice && String(ref.codice).trim()));
        // Ottieni le referenze per descrizione (opzionale): referenzeData e deviceRefData
        const referenzeDataLocal = JSON.parse(localStorage.getItem('referenzeData') || '[]');

        // Funzione per ottenere descrizione da referenzeData in base al codice
        function getDescrizioneForCode(cod) {
            for (const row of referenzeDataLocal) {
                const codeCell = row[0];
                const descCell = row[1];
                if (codeCell && String(codeCell).trim() === cod) {
                    return String(descCell || '').trim();
                }
            }
            return '';
        }

        // Funzione per ottenere cliente da deviceRefs
        function getClienteForCode(cod) {
            const found = deviceRefs.find(ref => ref.codice && String(ref.codice).trim() === cod);
            return found ? (found.cliente || '') : '';
        }

        // Utilità per convertire una lettera di colonna Excel in indice numerico
        const excelColToIndex = (colName) => {
            let index = 0;
            for (let i = 0; i < colName.length; i++) {
                index = index * 26 + (colName.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
            }
            return index - 1;
        };

        // Mappatura aggiornata delle colonne in base alla specifica dell'utente
        // Data -> colonna D (Domodossola)
        const colDate = excelColToIndex('D');
        // Codice -> colonna G (Genova)
        const colCode = excelColToIndex('G');
        // Descrizione -> colonna H (Hotel)
        const colDesc = excelColToIndex('H');
        // Cliente/Ragione sociale -> colonna S (Savona)
        const colCliente = excelColToIndex('S');
        // Lotto -> colonna O (Otranto)
        const colLotto = excelColToIndex('O');
        // Quantità -> colonna M (Milano)
        const colQuantity = excelColToIndex('M');
        // Unità di misura originale -> colonna N (Napoli) (non più utilizzata)
        const colUnit = excelColToIndex('N');
        // Flag scarti -> colonna I (indicatore "scarti" o quantità negativa)
        const colScartoFlag = excelColToIndex('I');

        const importRows = data.filter(row => row && row.length > Math.max(colUnit, colQuantity, colLotto, colCliente, colDesc));
        // Mappa per aggregare dati per codice, data e lotto
        const aggregated = {};
        importRows.forEach(row => {
            const code = String(row[colCode] || '').trim();
            if (!code || !validCodes.has(code)) {
                return; // ignora codici non presenti nel deviceRef
            }
            const rawDate = row[colDate];
            const dateStr = parseDateValue(rawDate);
            const lotto = String(row[colLotto] || '').trim();
            const key = `${code}||${dateStr}||${lotto}`;
            const qtyRaw = parseNumericValue(row[colQuantity]);
            const descrFlag = String(row[colScartoFlag] || '').toLowerCase();
            const descr = String(row[colDesc] || '').trim() || getDescrizioneForCode(code);
            const cliente = String(row[colCliente] || '').trim() || getClienteForCode(code);
            if (!aggregated[key]) {
                aggregated[key] = {
                    data: dateStr,
                    codice: code,
                    descrizione: descr,
                    cliente: cliente,
                    lotto: lotto,
                    quantita: 0,
                    unita: 0
                };
            }
            // Determina se è una riga di scarto
            const isScartoRow = (descrFlag.includes('scarto') || descrFlag.includes('scarti')) || (qtyRaw < 0);
            if (isScartoRow) {
                aggregated[key].unita += Math.abs(qtyRaw);
            } else {
                aggregated[key].quantita += qtyRaw;
            }
        });
        const finalData = Object.values(aggregated);
        localStorage.setItem('medicalProductionData', JSON.stringify(finalData));
        populateMedicalDeviceProductionTable(finalData);
    }

    /**
     * Popola la tabella di produzione medicale con i dati forniti.
     * Ogni riga contiene data, codice, descrizione, cliente, quantita, scarti e unita di misura.
     * @param {Array} data Array di oggetti di produzione medicale
     */
    function populateMedicalDeviceProductionTable(data) {
        if (!Array.isArray(data)) return;
        medicalDeviceTableBody.innerHTML = '';
        // Pre-carica i riferimenti DeviceRef una sola volta per evitare look-up ripetuti.
        let deviceRefs;
        try {
            // Usa la funzione getDeviceRefData se definita, altrimenti recupera dal localStorage.
            deviceRefs = typeof getDeviceRefData === 'function'
                ? getDeviceRefData()
                : JSON.parse(localStorage.getItem('deviceRefData') || '[]');
        } catch (e) {
            deviceRefs = [];
        }
        data.forEach(item => {
            const tr = document.createElement('tr');
            // Data
            const tdData = document.createElement('td');
            const dataInput = document.createElement('input');
            dataInput.type = 'text';
            dataInput.value = item.data || '';
            dataInput.readOnly = true;
            tdData.appendChild(dataInput);
            tr.appendChild(tdData);
            // Codice
            const tdCodice = document.createElement('td');
            const codiceInput = document.createElement('input');
            codiceInput.type = 'text';
            codiceInput.value = item.codice || '';
            codiceInput.readOnly = true;
            tdCodice.appendChild(codiceInput);
            tr.appendChild(tdCodice);
            // Descrizione
            const tdDesc = document.createElement('td');
            const descInput = document.createElement('input');
            descInput.type = 'text';
            descInput.value = item.descrizione || '';
            descInput.readOnly = true;
            tdDesc.appendChild(descInput);
            tr.appendChild(tdDesc);
            // Cliente
            const tdCliente = document.createElement('td');
            const clienteInput = document.createElement('input');
            clienteInput.type = 'text';
            clienteInput.value = item.cliente || '';
            clienteInput.readOnly = true;
            tdCliente.appendChild(clienteInput);
            tr.appendChild(tdCliente);
            // Lotto
            const tdLotto = document.createElement('td');
            const lottoInput = document.createElement('input');
            lottoInput.type = 'text';
            lottoInput.value = item.lotto || '';
            lottoInput.readOnly = true;
            tdLotto.appendChild(lottoInput);
            tr.appendChild(tdLotto);
            // Quantità (visualizzata come numero in pezzi, senza suffisso)
            const tdQty = document.createElement('td');
            const qtyInput = document.createElement('input');
            qtyInput.type = 'text';
            const qtyVal = parseFloat(item.quantita || 0);
            qtyInput.value = isNaN(qtyVal) ? '' : qtyVal;
            qtyInput.readOnly = true;
            tdQty.appendChild(qtyInput);
            tr.appendChild(tdQty);
            // Nuova colonna: numero teorico di scatoloni (arrotondato per eccesso)
            const tdBoxes = document.createElement('td');
            const boxesInput = document.createElement('input');
            boxesInput.type = 'text';
            boxesInput.readOnly = true;
            // Calcola i pezzi per scatolone a partire dai DeviceRef (colonna I o fallback)
            let perBox = NaN;
            try {
                const codeKey = String(item.codice || '').trim().toUpperCase();
                if (Array.isArray(deviceRefs)) {
                    const matchedRef = deviceRefs.find(ref => String(ref.codice || '').trim().toUpperCase() === codeKey);
                    if (matchedRef) {
                        const rawPerBox = (matchedRef.pezziPerScatolone !== undefined ? matchedRef.pezziPerScatolone : undefined) ||
                                          matchedRef.siringhePerScatola2 || matchedRef.siringhePerScatola;
                        if (rawPerBox != null && rawPerBox !== '') {
                            // Normalizza valori con separatori (es. "1.100" o "1,100")
                            const normalized = String(rawPerBox).replace(/\./g, '').replace(',', '.');
                            const n = parseFloat(normalized);
                            if (!isNaN(n) && n > 0) perBox = n;
                        }
                    }
                }
            } catch (e) {
                perBox = NaN;
            }
            let theoreticalBoxes = '';
            if (!isNaN(qtyVal) && qtyVal > 0 && !isNaN(perBox) && perBox > 0) {
                theoreticalBoxes = Math.ceil(qtyVal / perBox);
            }
            boxesInput.value = theoreticalBoxes;
            tdBoxes.appendChild(boxesInput);
            tr.appendChild(tdBoxes);
            medicalDeviceTableBody.appendChild(tr);
        });
    }
    referenzeInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            // 1. Sovrascrive i dati vecchi con quelli nuovi
            referenzeData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            // 2. Salva i nuovi dati e il nome del file nel browser per renderli persistenti
            saveStaticData('referenzeData', referenzeData, file.name);

            // 3. Aggiorna l'interfaccia mostrando il nome del nuovo file
            // Aggiorna l'etichetta con il nome del file ma non mostra più il flag visivo.
            referenzeFileStatusSpan.textContent = `File: ${file.name}`;
            referenzeFileStatusSpan.style.display = 'none';
            await showAlert('Nuovo file Referenze importato con successo!');

            // 4. Ricarica la logica di analisi e aggiorna la tabella per riflettere le modifiche
            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
            renderAnalysisTableHeaders();
            updateAnalisiTable();

        } catch (e) {
            await showAlert(`Errore durante l'importazione del file Referenze: ${e.message}`);
        }
    });

    pianoAnaliticoInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            // 1. Sovrascrive i dati vecchi con quelli nuovi
            pianoAnaliticoData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            // 2. Salva i nuovi dati e il nome del file nel browser
            saveStaticData('pianoAnaliticoData', pianoAnaliticoData, file.name);

            // 3. Aggiorna l'interfaccia
            // Aggiorna l'etichetta con il nome del file ma non mostra più il flag visivo.
            pianoAnaliticoFileStatusSpan.textContent = `File: ${file.name}`;
            pianoAnaliticoFileStatusSpan.style.display = 'none';
            await showAlert('Nuovo Piano Analitico importato con successo!');

            // 4. Ricarica la logica di analisi e aggiorna la tabella
            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
            renderAnalysisTableHeaders();
            updateAnalisiTable();

        } catch (e) {
            await showAlert(`Errore durante l'importazione del file Piano Analitico: ${e.message}`);
        }
    });

    // --- FINE BLOCCO DA INSERIRE ---

    pianoAnaliticoInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            // 1. Sovrascrive i dati vecchi con quelli nuovi
            pianoAnaliticoData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

            // 2. Salva i nuovi dati e il nome del file nel browser
            saveStaticData('pianoAnaliticoData', pianoAnaliticoData, file.name);

            // 3. Aggiorna l'interfaccia
            // Aggiorna l'etichetta con il nome del file ma non mostra più il flag visivo.
            pianoAnaliticoFileStatusSpan.textContent = `File: ${file.name}`;
            pianoAnaliticoFileStatusSpan.style.display = 'none';
            await showAlert('Nuovo Piano Analitico importato con successo!');

            // 4. Ricarica la logica di analisi e aggiorna la tabella
            loadAnalisiExcelData(pianoAnaliticoData, referenzeData);
            renderAnalysisTableHeaders();
            updateAnalisiTable();

        } catch (e) {
            await showAlert(`Errore durante l'importazione del file Piano Analitico: ${e.message}`);
        }
    });

    // --- FINE BLOCCO DA INSERIRE ---

    function loadLogbook() {
        try {
            const savedLog = localStorage.getItem('logbook_entries');
            if (savedLog) {
                logbookEntries = JSON.parse(savedLog);
                renderLogbook();
            }
        } catch (e) {
            console.error("Errore caricamento logbook:", e);
            logbookEntries = [];
        }
    }

    function saveLogbook() {
        try {
            localStorage.setItem('logbook_entries', JSON.stringify(logbookEntries));
        } catch (e) {
            console.error("Errore salvataggio logbook:", e);
        }
    }

// ========================================================================
    // ==> NUOVE FUNZIONI PER LA GESTIONE DEL FILE LAYOUT
    // ========================================================================
// ==> NUOVE FUNZIONI PER LA GESTIONE DEL FILE LAYOUT (CON PERSISTENZA)
// ========================================================================
let layoutData = {}; // Conterrà i dati del file Layout in formato {codice: layout}

// ===================================================================
    // ==> NUOVE VARIABILI PER LA TABELLA MEDICAL DEVICE <==
    // ===================================================================
    const medicalDeviceTableBody = document.querySelector('#medicalDeviceProductionTable tbody');
    const addMedicalDeviceRowBtn = document.getElementById('addMedicalDeviceRowBtn');
    const medicalDeviceStartDateInput = document.getElementById('medicalDeviceStartDate');
    const medicalDeviceEndDateInput = document.getElementById('medicalDeviceEndDate');
    const clearMedicalDeviceDateBtn = document.getElementById('clearMedicalDeviceDateBtn');
    const clearMedicalDeviceFiltersBtn = document.getElementById('clearMedicalDeviceFiltersBtn');
    const filterMedicalDeviceCodice = document.getElementById('filterMedicalDeviceCodice');
    const filterMedicalDeviceDescrizione = document.getElementById('filterMedicalDeviceDescrizione');
    const filterMedicalDeviceCliente = document.getElementById('filterMedicalDeviceCliente');
    // Nuovi campi di filtro per data e lotto nella tabella MD
    const filterMedicalDeviceData = document.getElementById('filterMedicalDeviceData');
    const filterMedicalDeviceLotto = document.getElementById('filterMedicalDeviceLotto');

    // Bottone per importazione produzione medical device
    const importMedicalProductionBtn = document.getElementById('importMedicalProductionBtn');
    if (importMedicalProductionBtn) {
        importMedicalProductionBtn.addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xls,.xlsx,.csv';
            input.addEventListener('change', async (e) => {
                const file = e.target.files[0];
                if (!file) return;
                await processMedicalProductionFile(file);
            });
            input.click();
        });
    }



// ===================================================================
    // ==> 1. FUNZIONE NUOVA DA AGGIUNGERE <==
    // Gestione cache dati statici (una volta al giorno)
    // ===================================================================
    /**
     * Controlla se è il primo avvio del giorno. Se lo è, invalida la cache
     * dei dati statici per incoraggiare la reimportazione di file aggiornati.
     */
    function gestisciCacheDatiStatici() {
        const oggi = new Date().toISOString().split('T')[0]; // Data in formato YYYY-MM-DD
        const ultimaVerifica = localStorage.getItem('staticCacheValidatedDate');

        if (ultimaVerifica !== oggi) {
            console.log("Primo avvio del giorno. Invalido la cache dei dati statici.");
            addLogEntry("Primo avvio del giorno: cache di Layout e Analisi resettata.");
            
            // Rimuove i dati vecchi dalla memoria del browser
            localStorage.removeItem('referenzeData');
            localStorage.removeItem('pianoAnaliticoData');
            localStorage.removeItem('layout_data');
            
            // Aggiorna la data dell'ultima verifica ad oggi
            localStorage.setItem('staticCacheValidatedDate', oggi);
            
            //showAlert("È il primo avvio di oggi. I dati di configurazione (Layout, Analisi) sono stati resettati. Per favore, reimporta le versioni più recenti se disponibili.");
        } else {
            console.log("Cache dati statici già validata per oggi.");
        }
    }

/**
 * Salva i dati del layout nel localStorage del browser per renderli persistenti.
 */
function saveLayoutData() {
    try {
        // Converte l'oggetto layoutData in una stringa JSON e lo salva
        localStorage.setItem('layout_data', JSON.stringify(layoutData));
        console.log("Dati Layout salvati nel localStorage.");
    } catch (e) {
        console.error("Errore nel salvataggio dei dati Layout:", e);
    }
}

function loadLayoutData() {
    try {
        const savedData = localStorage.getItem('layout_data');
        if (savedData) {
            layoutData = JSON.parse(savedData);
            
            // Mostra il flag se i dati sono già salvati (SENZA il nome del file)
            const layoutFlag = document.getElementById('layoutFileStatus');
            if (layoutFlag) {
                // Non mostrare più il flag visivo per il layout; verrà utilizzata soltanto la data di ultimo import.
                layoutFlag.style.display = 'none';
            }
            console.log("Dati Layout caricati correttamente dal localStorage.");
        }
    } catch (e) {
        console.error("Errore nel caricamento dei dati Layout:", e);
        layoutData = {};
    }
}

async function processLayoutFile(file) {
    addLogEntry(`--- Inizio importazione file Layout: ${file.name} ---`);
    try {
        const data = new Uint8Array(await file.arrayBuffer());
        const workbook = XLSX.read(data, {
            type: 'array'
        });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: false,
            defval: ""
        });

        const colIndex = {
            code: 1, // Colonna B
            family: 14, // Colonna O (Famiglia)
            layout: 37 // Colonna AL
        };

        const newLayoutData = {};
        jsonData.forEach((row, index) => {
            if (index === 0) return;
            const code = String(row[colIndex.code] || '').trim();
            const family = String(row[colIndex.family] || 'Senza Famiglia').trim(); // Legge la famiglia
            const layout = String(row[colIndex.layout] || '').trim();

            if (code) {
                // Memorizza un oggetto con entrambe le informazioni
                newLayoutData[code] = {
                    family: family,
                    layout: layout
                };
            }
        });

        layoutData = newLayoutData;
        localStorage.setItem('layout_fileName', file.name);

        saveLayoutData();

        const layoutFlag = document.getElementById('layoutFileStatus');
        if (layoutFlag) {
            // Non mostrare più il flag visivo per il layout; verrà utilizzata soltanto la data di ultimo import.
            layoutFlag.style.display = 'none';
        }
        await saveDataToServer();
        await showAlert(`File Layout "${file.name}" importato e salvato con successo.`);
        addLogEntry(`File Layout importato: ${Object.keys(layoutData).length} codici caricati.`);

        // Aggiorna la tabella degli arrivi con i nuovi dati di layout e famiglia
        document.querySelectorAll('#arrivalScheduleTable tbody tr').forEach(row => {
            const codeCell = row.cells[2].querySelector('input');
            const layoutCell = row.cells[4].querySelector('input');
            if (codeCell && layoutCell) {
                const layoutInfo = layoutData[codeCell.value] || { layout: '', family: 'Senza Famiglia' };
                layoutCell.value = layoutInfo.layout;
                row.dataset.family = layoutInfo.family; // Aggiorna anche l'attributo nascosto
            }
        });
        updateWarehouseGanttChart();

    } catch (error) {
        console.error("Errore durante l'importazione del file Layout:", error);
        addLogEntry(`Importazione Layout fallita: ${error.message}`);
        await showAlert(`Errore durante l'importazione del file Layout: ${error.message}.`);
    }
}
/**
 * Restituisce l'icona e il testo per il Gantt in base al valore del layout.
 */
function getLayoutIcon(layoutString) {
    if (!layoutString) return '';
    const layout = layoutString.toUpperCase();

    if (layout.includes('G5CELLA +20°C')) {
        return `<span class="layout-icon">(+20°C)</span>`;
    }
    if (layout.includes('G5CELLA +4°C')) {
        return `<span class="layout-icon"><span class="thermo-red">🌡️</span> (+4°C)</span>`;
    }
    if (layout.includes('G5CELLA -20°C')) {
        return `<span class="layout-icon"><span class="snow-blue">❄️</span> (-20°C)</span>`;
    }
    return '';
}

/**
 * Restituisce l'icona HTML corretta per un articolo in arrivo.
 * @param {object} task - L'oggetto contenente i dati della riga di arrivo.
 * @returns {string} La stringa HTML per l'icona.
 */
function getArrivalIconHtml(task) {
    const code = String(task.codiceArticolo || '').toUpperCase();
    const description = String(task.descrizioneArticolo || '').toLowerCase();

    // REGOLA ESTESA: Vale sia per PIL che per EGC
    if (code.startsWith('PIL') || code.startsWith('EGC')) {
        if (description.includes('etichetta')) {
            return '<span class="gantt-arrival-icon">🏷️</span>'; // Icona Etichetta
        }
        if (description.includes('astuccio')) {
            return '<span class="gantt-arrival-icon">📦</span>'; // Icona Astuccio
        }
    }

    // Mappa per i codici specifici che hai elencato
    const iconMap = {
        'PH701/50C': '💉', 'BEC0706': '💉', 'BEC0810': '💉', 'BEC1214': '💉', 'BEC1415': '💉',
        'BEC1818': '💉', 'BEC1918': '💉', 'BEC2420': '💉', 'BEC2520': '💉', 'BEC2620': '💉', 'BEC3024': '💉',
        'BEC0506': '<span class="md-icon">MD</span>', 'BEC0910': '<span class="md-icon">MD</span>',
        'BEC1010': '<span class="md-icon">MD</span>', 'BEC1111': '<span class="md-icon">MD</span>',
        'BOR1321': '<span class="md-icon">MD</span>',
        'STV0214': '💉', 'STV0314': '💉', 'STV1014': '💉', 'STV1114': '💉', 'STV2420': '💉',
        'CAP0106': '💉', 'CAP0308': '💉', 'CAP0408': '💉', 'CAP0513': '💉', 'CAP0613': '💉',
        // Aggiungi qui altri codici se necessario
    };
    
    // Se il codice è nella mappa, restituisce l'icona corrispondente
    if (iconMap[code]) {
        return iconMap[code].startsWith('<span') 
             ? iconMap[code] 
             : `<span class="gantt-arrival-icon">${iconMap[code]}</span>`;
    }

    return ''; // Nessuna icona se non corrisponde a nessuna regola
}


    /**
     * Restituisce l'icona e il testo per il Gantt in base al valore del layout.
     */
    function getLayoutIcon(layoutString) {
        if (!layoutString) return '';
        const layout = layoutString.toUpperCase();

        if (layout.includes('G5CELLA +20°C')) {
            return `<span class="layout-icon">(+20°C)</span>`;
        }
        if (layout.includes('G5CELLA +4°C')) {
            return `<span class="layout-icon"><span class="thermo-red">🌡️</span> (+4°C)</span>`;
        }
        if (layout.includes('G5CELLA -20°C')) {
            return `<span class="layout-icon"><span class="snow-blue">❄️</span> (-20°C)</span>`;
        }
        return '';
    }

// ========================================================================
// ==> NUOVE FUNZIONI SPECIFICHE PER LA TABELLA ARRIVI
// ========================================================================

/**
 * Crea una riga per la tabella degli arrivi.
 */


function createArrivalScheduleRow(rowData = {}) {
    const row = document.createElement('tr');

    row.dataset.family = rowData.family || 'Senza Famiglia';
    // Memorizza lo stato del magazzino (merda da evadere / evasa) nel dataset.  Se non
    // specificato nei dati, imposta lo stato predefinito a bianco (da evadere).
    row.dataset.magStatus = rowData.magStatus || 'white';

    if (isMedicalDeviceCode(rowData.codiceArticolo)) {
        row.classList.add('production-4xxxx-bg');
    }

    // Funzione di supporto per gestire le virgolette
    const escapeAttr = (str) => String(str || '').replace(/"/g, '&quot;');

    row.innerHTML = `
        <td><input type="checkbox" class="arrival-row-selector"></td>
        <td><input type="text" value="${escapeAttr(rowData.ov || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.codiceArticolo || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.descrizioneArticolo || '')}" style="text-align: left;"></td>
        <td><input type="text" value="${escapeAttr(rowData.layout || '')}" readonly></td>
        <td><input type="number" value="${escapeAttr(rowData.quantita || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.um || '')}"></td>
        <td><input type="text" class="datepicker" value="${escapeAttr(rowData.dataConsegna || '')}"></td>
        <td><input type="text" class="datepicker" value="${escapeAttr(rowData.dataConferma || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.ragioneSociale || '')}" style="text-align: left;"></td>
        <td><input type="text" value="${escapeAttr(rowData.riferimentoCliente || '')}" style="text-align: left;"></td>
        <td><input type="text" value="${escapeAttr(rowData.indirizzo || '')}" style="text-align: left;"></td>
        <td><input type="text" value="${escapeAttr(rowData.cap || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.citta || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.provincia || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.telefono || '')}"></td>
    `;

    row.querySelectorAll('.datepicker').forEach(input => {
        flatpickr(input, {
            dateFormat: "d/m/Y",
            locale: "it"
        });
    });

    // Memorizza le note di servizio (colonna Q) nel dataset della riga, così da
    // poterle recuperare successivamente nel tooltip senza aggiungere una colonna visibile
    row.dataset.noteServizio = rowData.noteServizio || '';


    row.querySelectorAll('input').forEach(input => {
        input.addEventListener('change', () => {
            updateWarehouseGanttChart();
            autoSaveAllData();
        });
    });

    return row;
}
function getArrivalScheduleRowData(row) {
    const cells = row.cells;
    return {
        ov: cells[1].querySelector('input').value,
        codiceArticolo: cells[2].querySelector('input').value,
        descrizioneArticolo: cells[3].querySelector('input').value,
        layout: cells[4].querySelector('input').value,
        quantita: cells[5].querySelector('input').value,
        um: cells[6].querySelector('input').value,
        dataConsegna: cells[7].querySelector('input').value,
        dataConferma: cells[8].querySelector('input').value,
        ragioneSociale: cells[9].querySelector('input').value,
        riferimentoCliente: cells[10].querySelector('input').value,
        indirizzo: cells[11].querySelector('input').value,
        cap: cells[12].querySelector('input').value,
        citta: cells[13].querySelector('input').value,
        provincia: cells[14].querySelector('input').value,
        telefono: cells[15].querySelector('input').value,
        family: row.dataset.family || 'Senza Famiglia', // Legge la famiglia dall'attributo dati
        // Ritorna anche le note di servizio memorizzate nel dataset (non visibili in tabella)
        noteServizio: row.dataset.noteServizio || '',
        // Stato magazzino: "white" (da evadere) oppure "green" (evasa)
        magStatus: row.dataset.magStatus || 'white'
    };
}

function getAllArrivalData() {
    const data = [];
    document.querySelectorAll('#arrivalScheduleTable tbody tr').forEach(row => {
        data.push(getArrivalScheduleRowData(row));
    });
    return data;
}

// ===================================================================
    // ==> NUOVE FUNZIONI PER GESTIRE LA TABELLA "MERCE NON ARRIVATA" <==
    // ===================================================================

    /**
     * Crea una riga di sola lettura per la tabella della merce non arrivata.
     */
    function createOverdueArrivalRow(rowData = {}) {
        const row = createArrivalScheduleRow(rowData); // Riusa la funzione esistente
        // Rende tutti gli input nella riga non modificabili
        row.querySelectorAll('input').forEach(input => {
            input.readOnly = true;
            input.style.cursor = 'not-allowed';
        });
        // Disabilita la checkbox
        const checkbox = row.querySelector('.arrival-row-selector');
        if (checkbox) checkbox.disabled = true;
        
        return row;
    }

    function populateOverdueTable(overdueItems) {
    if (!overdueArrivalsTableBody) return;
    overdueArrivalsTableBody.innerHTML = '';

    // Ordina dalla data più recente alla più vecchia
    overdueItems.sort((a, b) => {
        const parse = (str) => {
            if (!str) return new Date(0);
            const p = str.split('/');
            return new Date(p[2], p[1] - 1, p[0]);
        };
        return parse(b.dataConsegna) - parse(a.dataConsegna);
    });

    overdueItems.forEach(item => {
        overdueArrivalsTableBody.appendChild(createOverdueArrivalRow(item));
    });
    // Dopo l'import, aggiorna i filtri (ad es. se sono attivi)
    applyOverdueFilters && applyOverdueFilters();
}



function applyOverdueFilters() {
    const ov = document.getElementById('filterOverdueOV').value.trim().toLowerCase();
    const codice = document.getElementById('filterOverdueCodice').value.trim().toLowerCase();
    const descr = document.getElementById('filterOverdueDescrizione').value.trim().toLowerCase();
    const ragSoc = document.getElementById('filterOverdueRagSoc').value.trim().toLowerCase();
    const dataDa = document.getElementById('filterOverdueDataDa').value;
    const dataA = document.getElementById('filterOverdueDataA').value;

    // Converti le date gg/mm/aaaa in oggetti Date
    const parseDate = (str) => {
        if (!str) return null;
        const parts = str.split('/');
        if (parts.length !== 3) return null;
        return new Date(parts[2], parts[1] - 1, parts[0]);
    };
    const dateFrom = parseDate(dataDa);
    const dateTo = parseDate(dataA);

    document.querySelectorAll('#overdueArrivalsTable tbody tr').forEach(row => {
        const cells = row.querySelectorAll('td');
        const valOV = cells[1].querySelector('input').value.trim().toLowerCase();
        const valCodice = cells[2].querySelector('input').value.trim().toLowerCase();
        const valDescr = cells[3].querySelector('input').value.trim().toLowerCase();
        const valData = cells[7].querySelector('input').value.trim();
        const valRagSoc = cells[9].querySelector('input').value.trim().toLowerCase();

        let show = true;

        if (ov && !valOV.includes(ov)) show = false;
        if (codice && !valCodice.includes(codice)) show = false;
        if (descr && !valDescr.includes(descr)) show = false;
        if (ragSoc && !valRagSoc.includes(ragSoc)) show = false;

        if ((dateFrom || dateTo) && valData) {
            const parts = valData.split('/');
            if (parts.length === 3) {
                const rowDate = new Date(parts[2], parts[1] - 1, parts[0]);
                if (dateFrom && rowDate < dateFrom) show = false;
                if (dateTo && rowDate > dateTo) show = false;
            }
        }

        row.style.display = show ? '' : 'none';
    });
}

// Attiva i filtri live
['filterOverdueOV','filterOverdueCodice','filterOverdueDescrizione','filterOverdueRagSoc','filterOverdueDataDa','filterOverdueDataA'].forEach(id => {
    document.getElementById(id).addEventListener('input', applyOverdueFilters);
});
document.getElementById('clearOverdueFiltersBtn').addEventListener('click', function() {
    ['filterOverdueOV','filterOverdueCodice','filterOverdueDescrizione','filterOverdueRagSoc','filterOverdueDataDa','filterOverdueDataA'].forEach(id => {
        document.getElementById(id).value = '';
    });
    applyOverdueFilters();
});
// Inizializza flatpickr per i due campi data
flatpickr(document.getElementById('filterOverdueDataDa'), { dateFormat: "d/m/Y", locale: "it" });
flatpickr(document.getElementById('filterOverdueDataA'), { dateFormat: "d/m/Y", locale: "it" });

    /**
     * Filtra le righe della tabella "Merce in Quarantena" in base ai valori
     * inseriti nei campi di filtro.  La logica è analoga a quella usata
     * per la tabella "Merce non Arrivata": confronta OV, codice articolo,
     * descrizione, ragione sociale e date di consegna.  I campi filtro
     * sono definiti nella sezione HTML di quarantena.
     */
    function applyQuarantineFilters() {
        const ovVal = document.getElementById('filterQuarantineOV').value.trim().toLowerCase();
        const codiceVal = document.getElementById('filterQuarantineCodice').value.trim().toLowerCase();
        const descrVal = document.getElementById('filterQuarantineDescrizione').value.trim().toLowerCase();
        const ragSocVal = document.getElementById('filterQuarantineRagSoc').value.trim().toLowerCase();
        const dataDaVal = document.getElementById('filterQuarantineDataDa').value;
        const dataAVal = document.getElementById('filterQuarantineDataA').value;
        // Parsing date helper
        const parseDate = (str) => {
            if (!str) return null;
            const parts = str.split('/');
            if (parts.length !== 3) return null;
            return new Date(parts[2], parts[1] - 1, parts[0]);
        };
        const dateFrom = parseDate(dataDaVal);
        const dateTo = parseDate(dataAVal);
        document.querySelectorAll('#quarantineTable tbody tr').forEach(row => {
            const cells = row.querySelectorAll('td');
            // Le colonne sono identiche a quelle degli arrivi: indice 1 OV, 2 Codice, 3 Descrizione, 7 Data Consegna, 9 Ragione Sociale
            const valOV = cells[1].querySelector('input').value.trim().toLowerCase();
            const valCodice = cells[2].querySelector('input').value.trim().toLowerCase();
            const valDescr = cells[3].querySelector('input').value.trim().toLowerCase();
            const valLayout = cells[4].querySelector('input').value.trim().toLowerCase();
            const valData = cells[7].querySelector('input').value.trim();
            const valRagSoc = cells[9].querySelector('input').value.trim().toLowerCase();
            let show = true;
            if (ovVal && !valOV.includes(ovVal)) show = false;
            if (codiceVal && !valCodice.includes(codiceVal)) show = false;
            if (descrVal && !valDescr.includes(descrVal)) show = false;
            if (ragSocVal && !valRagSoc.includes(ragSocVal)) show = false;
            if ((dateFrom || dateTo) && valData) {
                const parts = valData.split('/');
                if (parts.length === 3) {
                    const rowDate = new Date(parts[2], parts[1] - 1, parts[0]);
                    if (dateFrom && rowDate < dateFrom) show = false;
                    if (dateTo && rowDate > dateTo) show = false;
                }
            }
            row.style.display = show ? '' : 'none';
        });
    }

    // Attiva i filtri live per la tabella quarantena
    ['filterQuarantineOV','filterQuarantineCodice','filterQuarantineDescrizione','filterQuarantineRagSoc','filterQuarantineDataDa','filterQuarantineDataA'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', applyQuarantineFilters);
    });
    // Gestisce il reset dei filtri per la tabella quarantena
    const clearQuarantineFiltersBtn = document.getElementById('clearQuarantineFiltersBtn');
    if (clearQuarantineFiltersBtn) {
        clearQuarantineFiltersBtn.addEventListener('click', () => {
            ['filterQuarantineOV','filterQuarantineCodice','filterQuarantineDescrizione','filterQuarantineRagSoc','filterQuarantineDataDa','filterQuarantineDataA'].forEach(id => {
                const input = document.getElementById(id);
                if (input) input.value = '';
            });
            applyQuarantineFilters();
        });
    }
    // Inizializza flatpickr per i campi data nella tabella quarantena
    flatpickr(document.getElementById('filterQuarantineDataDa'), { dateFormat: "d/m/Y", locale: "it" });
    flatpickr(document.getElementById('filterQuarantineDataA'), { dateFormat: "d/m/Y", locale: "it" });

    /**
     * Raccoglie i dati da una singola riga della tabella merce non arrivata.
     */
    function getOverdueArrivalRowData(row) {
        // La struttura è identica, quindi possiamo riutilizzare la funzione esistente
        return getArrivalScheduleRowData(row); 
    }

    /**
     * Raccoglie tutti i dati dalla tabella merce non arrivata.
     */
    function getAllOverdueArrivalData() {
        const data = [];
        if (overdueArrivalsTableBody) {
            overdueArrivalsTableBody.querySelectorAll('tr').forEach(row => {
                data.push(getOverdueArrivalRowData(row));
            });
        }
        return data;
    }


function applyOpiFilters() {
    const opVal = document.getElementById('filterOpiOP').value.trim().toLowerCase();
    const ovVal = document.getElementById('filterOpiOV').value.trim().toLowerCase();
    const codiceVal = document.getElementById('filterOpiCodice').value.trim().toLowerCase();
    const articoloVal = document.getElementById('filterOpiArticolo').value.trim().toLowerCase();
    const clienteVal = document.getElementById('filterOpiCliente').value.trim().toLowerCase();
    const lottoVal = document.getElementById('filterOpiLotto').value.trim().toLowerCase();
    const quantitaVal = document.getElementById('filterOpiQuantita').value.trim().toLowerCase();
    const umVal = document.getElementById('filterOpiUM').value.trim().toLowerCase();
    const operatoreVal = document.getElementById('filterOpiOperatore').value.trim().toLowerCase();

    // Date
    const dataProdDa = document.getElementById('opiStartDate').value;
    const dataProdA = document.getElementById('opiEndDate').value;
    const scadDa = document.getElementById('opiScadStartDate').value;
    const scadA = document.getElementById('opiScadEndDate').value;

    // Funzione di parsing
    function parseDate(str) {
        if (!str) return null;
        const p = str.split('/');
        if (p.length !== 3) return null;
        return new Date(p[2], p[1] - 1, p[0]);
    }
    const dataProdDaDate = parseDate(dataProdDa);
    const dataProdADate = parseDate(dataProdA);
    const scadDaDate = parseDate(scadDa);
    const scadADate = parseDate(scadA);

    document.querySelectorAll('#opiTable tbody tr').forEach(row => {
        const cells = row.querySelectorAll('td');
        let show = true;
        if (opVal && !cells[1].textContent.toLowerCase().includes(opVal)) show = false;
        if (ovVal && !cells[2].textContent.toLowerCase().includes(ovVal)) show = false;
        if (codiceVal && !cells[3].textContent.toLowerCase().includes(codiceVal)) show = false;
        if (articoloVal && !cells[4].textContent.toLowerCase().includes(articoloVal)) show = false;
        if (clienteVal && !cells[5].textContent.toLowerCase().includes(clienteVal)) show = false;
        if (lottoVal && !cells[6].textContent.toLowerCase().includes(lottoVal)) show = false;
        if (quantitaVal && !cells[7].textContent.toLowerCase().includes(quantitaVal)) show = false;
        if (umVal && !cells[8].textContent.toLowerCase().includes(umVal)) show = false;
        if (operatoreVal && !cells[9].textContent.toLowerCase().includes(operatoreVal)) show = false;

        // Filtri data produzione
        const dataProdStr = cells[0].textContent.trim();
        const dataProdDate = parseDate(dataProdStr);
        if (show && dataProdDaDate && (!dataProdDate || dataProdDate < dataProdDaDate)) show = false;
        if (show && dataProdADate && (!dataProdDate || dataProdDate > dataProdADate)) show = false;

        // Filtri scadenza lotto
        const scadStr = cells[10].textContent.trim();
        const scadDate = parseDate(scadStr);
        if (show && scadDaDate && (!scadDate || scadDate < scadDaDate)) show = false;
        if (show && scadADate && (!scadDate || scadDate > scadADate)) show = false;

        row.style.display = show ? '' : 'none';
    });
}

[
    'filterOpiOP', 'filterOpiOV', 'filterOpiCodice', 'filterOpiArticolo', 'filterOpiCliente',
    'filterOpiLotto', 'filterOpiQuantita', 'filterOpiUM', 'filterOpiOperatore',
    'opiStartDate', 'opiEndDate', 'opiScadStartDate', 'opiScadEndDate'
].forEach(id => {
    document.getElementById(id).addEventListener('input', applyOpiFilters);
    document.getElementById(id).addEventListener('change', applyOpiFilters);
});

// Reset filtri TUTTO (inclusi calendari)
document.getElementById('clearOpiFiltersBtn').addEventListener('click', function() {
    [
        'filterOpiOP','filterOpiOV','filterOpiCodice','filterOpiArticolo',
        'filterOpiCliente','filterOpiLotto','filterOpiQuantita','filterOpiUM','filterOpiOperatore',
        'opiStartDate','opiEndDate','opiScadStartDate','opiScadEndDate'
    ].forEach(id => document.getElementById(id).value = '');
    applyOpiFilters();
});

// Attivazione filtri live (se non già presente)
[
    'filterOpiOP','filterOpiOV','filterOpiCodice','filterOpiArticolo',
    'filterOpiCliente','filterOpiLotto','filterOpiQuantita','filterOpiUM','filterOpiOperatore'
].forEach(id => {
    document.getElementById(id).addEventListener('input', applyOpiFilters);
});



function exportPropostaLayoutPDF() {
    const allArrivalRows = getAllArrivalData();
    if (allArrivalRows.length === 0) {
        showAlert('Nessun dato presente nella tabella degli arrivi da esportare.');
        return;
    }

    // Nuova logica: filtra la proposta di layout in base alle date scelte dall'utente nel filtro arrivi.
    // Se non vengono specificate date, usa l'intervallo di default (oggi + 14 giorni)
    let startDate = null;
    let endDate = null;
    const arrivalStartDateInput = document.getElementById('arrivalStartDate');
    const arrivalEndDateInput = document.getElementById('arrivalEndDate');
    if (arrivalStartDateInput && arrivalStartDateInput.value) {
        startDate = flatpickr.parseDate(arrivalStartDateInput.value, "d/m/Y");
    }
    if (arrivalEndDateInput && arrivalEndDateInput.value) {
        endDate = flatpickr.parseDate(arrivalEndDateInput.value, "d/m/Y");
        if (endDate) endDate.setHours(23, 59, 59, 999);
    }
    // Se non specificati, default a oggi e prossimi 14 giorni
    if (!startDate) {
        startDate = new Date();
        startDate.setHours(0, 0, 0, 0);
    }
    if (!endDate) {
        endDate = new Date(startDate.getTime());
        endDate.setDate(startDate.getDate() + 14);
        endDate.setHours(23, 59, 59, 999);
    }
    let filteredData = allArrivalRows.filter(row => {
        const parts = (row.dataConsegna || '').split('/');
        if (parts.length !== 3) return false;
        const rowDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        return (!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate);
    });
    if (filteredData.length === 0) {
        const msgRange = arrivalStartDateInput && arrivalStartDateInput.value ?
            `nessun arrivo previsto nel periodo selezionato` :
            'Nessun arrivo previsto nei prossimi 14 giorni da includere nella proposta.';
        showAlert(msgRange.charAt(0).toUpperCase() + msgRange.slice(1));
        return;
    }

    // NUOVA LOGICA DI ORDINAMENTO: Prima per famiglia, poi per data di consegna
    filteredData.sort((a, b) => {
        // 1. Ordina per nome della famiglia (alfabeticamente)
        const familyCompare = (a.family || 'Senza Famiglia').localeCompare(b.family || 'Senza Famiglia');
        if (familyCompare !== 0) {
            return familyCompare;
        }

        // 2. Se le famiglie sono uguali, ordina per data
        const dateAParts = a.dataConsegna.split('/');
        const dateBParts = b.dataConsegna.split('/');
        const dateA = new Date(parseInt(dateAParts[2]), parseInt(dateAParts[1]) - 1, parseInt(dateAParts[0]));
        const dateB = new Date(parseInt(dateBParts[2]), parseInt(dateBParts[1]) - 1, parseInt(dateBParts[0]));
        return dateA - dateB;
    });

    // NUOVA LOGICA DI RENDERIZZAZIONE: Cicla i dati ordinati e aggiunge le intestazioni
    let tableRowsHtml = '';
    let currentFamily = null;

    filteredData.forEach(row => {
        // Controlla se la famiglia è cambiata rispetto alla riga precedente
        if (row.family !== currentFamily) {
            currentFamily = row.family;
            // Aggiunge la riga di intestazione in grassetto per la nuova famiglia
            tableRowsHtml += `
                <tr>
                    <td colspan="9" style="font-weight: bold; text-align: left; background-color: #e0e0e0; padding: 6px 10px; font-size: 1.1em;">
                        ${currentFamily || 'Senza Famiglia'}
                    </td>
                </tr>
            `;
        }

        // Aggiunge la riga con i dati dell'articolo (logica precedente)
        tableRowsHtml += `
            <tr>
                <td>${row.ov || ''}</td>
                <td>${row.dataConsegna || ''}</td>
                <td>${row.codiceArticolo || ''}</td>
                <td class="desc-cell">${row.descrizioneArticolo || ''}</td>
                <td>${row.quantita || ''}</td>
                <td>${row.um || ''}</td>
                <td style="font-weight: bold;">${row.layout || 'N/D'}</td>
                <td class="confirmation-cell">
                    <span class="choice-box yes">Sì</span>
                    <span class="choice-box no">No</span>
                </td>
                <td class="real-layout-cell"></td>
            </tr>
        `;
    });

    // Prepara la stringa che descrive il periodo usato nel filtro
    let periodText = '';
    if (arrivalStartDateInput && arrivalStartDateInput.value) {
        periodText = `Dal ${arrivalStartDateInput.value}`;
    }
    if (arrivalEndDateInput && arrivalEndDateInput.value) {
        periodText += periodText ? ` al ${arrivalEndDateInput.value}` : `Al ${arrivalEndDateInput.value}`;
    }
    if (!periodText) {
        // Se non specificato, mostra il periodo predefinito utilizzato (startDate e endDate calcolati)
        const startStr = startDate ? startDate.toLocaleDateString('it-IT') : '';
        const endStr = endDate ? endDate.toLocaleDateString('it-IT') : '';
        periodText = startStr && endStr ? `Dal ${startStr} al ${endStr}` : '';
    }
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
        <head>
            <title>Proposta Layout Magazzino</title>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Quicksand:wght@400;500;700&display=swap');
                body { font-family: 'Quicksand', sans-serif; margin: 1.5cm; }
                h1 { color: #2c3e50; text-align: center; border-bottom: 2px solid #ccc; padding-bottom: 10px; }
                p { text-align: center; font-size: 0.9em; color: #555; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 9pt; table-layout: fixed; }
                th, td { border: 1px solid #999; padding: 8px 5px; text-align: center; vertical-align: middle; word-wrap: break-word; }
                th { background-color: #f2f2f2; color: #333; }
                .desc-cell { text-align: left; font-size: 0.85em; }
                .confirmation-cell { display: flex; justify-content: space-around; align-items: center; border: none; padding: 5px; }
                .choice-box { display: inline-block; width: 35px; padding: 5px; border: 1px solid #aaa; border-radius: 4px; text-align: center; }
                .choice-box.yes { background-color: #E8F5E9; color: #1B5E20; border-color: #81C784;}
                .choice-box.no { background-color: #FFEBEE; color: #B71C1C; border-color: #EF9A9A;}
                .real-layout-cell { height: 30px; }
                @page { size: A4 landscape; margin: 1cm; }
            </style>
        </head>
        <body>
            <h1>Proposta Layout Magazzino - Merce in Arrivo</h1>
            <p>Data di stampa: ${new Date().toLocaleDateString('it-IT')}</p>
            ${periodText ? `<p>Periodo: ${periodText}</p>` : ''}
            <table>
                <thead>
                    <tr>
                        <th style="width: 6%;">OV</th>
                        <th style="width: 8%;">Data Consegna</th>
                        <th style="width: 10%;">Codice Articolo</th>
                        <th style="width: 33%; text-align: left;">Descrizione Articolo</th>
                        <th style="width: 6%;">Quantità</th>
                        <th style="width: 4%;">UM</th>
                        <th style="width: 12%;">Layout Consigliato</th>
                        <th style="width: 11%;">Confermato?</th>
                        <th style="width: 10%;">Layout Reale</th>
                    </tr>
                </thead>
                <tbody>
                    ${tableRowsHtml}
                </tbody>
            </table>
        </body>
        </html>
    `);
    printWindow.document.close();
    printWindow.focus();
    setTimeout(() => {
        printWindow.print();
    }, 500);
}

    function addLogEntry(message) {
        const timestamp = new Date().toLocaleString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });

        let cleanMessage = message;
        if (cleanMessage.includes("da Importazione PP da Excel (manuale)")) {
            cleanMessage = cleanMessage.replace(" da Importazione PP da Excel (manuale)", "");
        } else if (cleanMessage.includes("da Download automatico")) {
            cleanMessage = cleanMessage.replace(" da Download automatico", "");
        }
        if (cleanMessage.includes("da Importazione OV")) {
            cleanMessage = cleanMessage.replace(" da Importazione OV", "");
        }

        logbookEntries.unshift(`${timestamp}: ${cleanMessage}`);
        const maxLogEntries = 200;
        if (logbookEntries.length > maxLogEntries) {
            logbookEntries = logbookEntries.slice(0, maxLogEntries);
        }
        renderLogbook();
        saveLogbook();
    }

    function renderLogbook(entries = logbookEntries) {
        if (logbookContentElement) {
            logbookContentElement.textContent = entries.join('\n');
        }
    }

    function compareAndApplyChanges(newData, source) {
        const oldData = getAllTableData();
        const oldDataMap = new Map(oldData.map(row => [row.codice, row]));
        const newDataMap = new Map(newData.map(row => [row.codice, row]));
        let changesFound = false;

        newDataMap.forEach((newRowData, codice) => {
            const oldRowData = oldDataMap.get(codice);
            if (!oldRowData) {
                addLogEntry(`Aggiunta riga (${newRowData.codice} - ${newRowData.prodotto})`);
                productionTableBody.appendChild(createRow(newRowData));
                changesFound = true;
            } else {
                const rowElement = Array.from(productionTableBody.querySelectorAll('.code-input')).find(input => input.value === codice)?.closest('tr');
                if(rowElement) {
                    const fieldsToCompare = ['prodotto', 'cliente', 'quantitaRichiesta', 'giacenzaMagazzino', 'quantitaDaProdurre', 'produzioneData', 'dataConfezionamento', 'dataSpedizione', 'lottoSC', 'note', 'macchinari', 'operatore'];
                    fieldsToCompare.forEach(field => {
                        const oldValue = oldRowData[field] || "";
                        const newValue = newRowData[field] || "";
                        if (String(oldValue).trim() !== String(newValue).trim()) {
                            addLogEntry(`Modifica riga (${codice} - ${newRowData.prodotto}): campo '${field}' da '${oldValue}' a '${newValue}'${source ? ` (da ${source})` : ''}`);
                            const input = rowElement.querySelector(`.col-${field.replace(/([A-Z])/g, '-$1').toLowerCase()} input, .col-${field.replace(/([A-Z])/g, '-$1').toLowerCase()} select`);
                            if(input) {
                                input.value = newValue;
                                if (input.classList.contains('si-no-select')) {
                                    setupSiNoSelect(input);
                                }
                            }
                            changesFound = true;
                        }
                    });
                }
                oldDataMap.delete(codice);
            }
        });

        oldDataMap.forEach((oldRowData, codice) => {
            addLogEntry(`Eliminata riga (${oldRowData.codice} - ${oldRowData.prodotto})${source ? ` (da ${source})` : ''}`);
            const rowElement = Array.from(productionTableBody.querySelectorAll('.code-input')).find(input => input.value === codice)?.closest('tr');
            if (rowElement) rowElement.remove();
            changesFound = true;
        });

        if (changesFound) {
            addLogEntry(`${source}: Modifiche applicate con successo.`);
            productionTableBody.querySelectorAll('tr').forEach(row => validateRow(row));
            updateGanttChart();
            updateWarehouseGanttChart();
            updateDailyProductionTable();
            updateAnalisiTable();
            runFullCheck();
            autoSaveAllData();
        } else {
            addLogEntry(`${source}: Nessuna modifica trovata.`);
        }
    }
    async function autoDownloadAndImport() {
        const now = new Date();
        const currentTimeInMinutes = now.getHours() * 60 + now.getMinutes();

        if (currentTimeInMinutes < (7 * 60 + 30) || currentTimeInMinutes > (18 * 60)) {
            console.log("Download automatico saltato: fuori dall'orario di lavoro.");
            return;
        }

        const monthNames = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
        const year = now.getFullYear();
        const month = monthNames[now.getMonth()];
        const day = String(now.getDate()).padStart(2, '0');
        const monthNum = String(now.getMonth() + 1).padStart(2, '0');
        const yearShort = String(year).slice(-2);

        const datePart = `${day}.${monthNum}.${yearShort}`;

        const baseFileNameWithSpace = `programma produzione ${datePart}`;
        const baseFileNameWithoutSpace = `programmaproduzione ${datePart}`;

        const pathsToTry = [
            `\\\SRVFS\\Dati\\Produzione\\programma prod\\${year}\\${month}\\${baseFileNameWithSpace}.xlsx`,
            `\\\SRVFS\\Dati\\Produzione\\programma prod\\${year}\\${month}\\${baseFileNameWithSpace}.xls`,
            `\\\SRVFS\\Dati\\Produzione\\programma prod\\${year}\\${month}\\${baseFileNameWithoutSpace}.xlsx`,
            `\\\SRVFS\\Dati\\Produzione\\programma prod\\${year}\\${month}\\${baseFileNameWithoutSpace}.xls`
        ];

        let fileFoundAndProcessed = false;
        for (const path of pathsToTry) {
            addLogEntry(`Avvio tentativo download automatico da: ${path}`);

            const simulateFileExists = false;
            if (simulateFileExists) {
                fileFoundAndProcessed = true;
                break;
            }
        }

        if (!fileFoundAndProcessed) {
            addLogEntry("Download automatico: file non trovato nel percorso specifico (controllate varianti con/senza spazio e .xlsx/.xls).");
        }
    }

    // Cerca la funzione esistente e sostituiscila con questa
    function startAutoImportTimer() {

        // Funzione principale che gestisce il processo di importazione
        async function autoImportPP() {
            const now = new Date();
            const currentHour = now.getHours();
            const currentMinutes = now.getMinutes();
            const currentTime = currentHour + currentMinutes / 60;

            // 1. Controlla se siamo nell'intervallo di tempo attivo (7:30 - 18:00)
            if (currentTime < 7.5 || currentTime > 18.0) {
                console.log("Auto-import saltato: fuori dall'orario di lavoro.");
                return;
            }

            // 2. Costruisce dinamicamente i nomi del mese e del file
            const year = now.getFullYear();
            const monthNames = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
            const month = monthNames[now.getMonth()];
            const day = String(now.getDate()).padStart(2, '0');
            const monthNum = String(now.getMonth() + 1).padStart(2, '0');
            const yearShort = String(year).slice(-2);
            const dateForFileName = `${day}.${monthNum}.${yearShort}`;
            const fileNameBase = `programma produzione ${dateForFileName}`;

            // 3. Prepara la lista dei percorsi da tentare in ordine di priorità
            const pathsToTry = [
                `file:///SRVFS/Dati/Produzione/programma prod/${year}/${month}/${fileNameBase}.xlsx`,
                `file:///SRVFS/Dati/Produzione/programma prod/${year}/${month}/${fileNameBase}.xls`, // Aggiunto fallback per .xls
                `file:///SRVFS/Dati/Produzione/programma prod/${year}/${fileNameBase}.xlsx`,
                `file:///SRVFS/Dati/Produzione/programma prod/${year}/${fileNameBase}.xls`,
                `file:///SRVFS/Dati/Produzione/programma prod/${fileNameBase}.xlsx`,
                `file:///SRVFS/Dati/Produzione/programma prod/${fileNameBase}.xls`
            ];

            for (const path of pathsToTry) {
                try {
                    addLogEntry(`Tentativo di importazione automatica da: ${path}`);
                    const response = await fetch(path);

                    if (response.ok) {
                        const arrayBuffer = await response.arrayBuffer();

                        // Simula un oggetto 'File' per la funzione processPPFile
                        const simulatedFile = {
                            name: path.split('/').pop(),
                            arrayBuffer: () => Promise.resolve(arrayBuffer)
                        };

                        // 4. Passa il file alla funzione di processamento esistente
                        await processPPFile(simulatedFile);
                        addLogEntry(`File importato con successo da: ${path}`);
                        return; // Interrompe il ciclo se il file viene trovato e processato
                    }
                } catch (error) {
                    // 5. Errore gestito silenziosamente, come richiesto
                    console.warn(`Download fallito da ${path}:`, error);
                    addLogEntry(`Download da ${path} fallito.`);
                }
            }
        }

        // Esegui il primo tentativo 5 secondi dopo l'avvio dell'applicazione
        setTimeout(autoImportPP, 5000);

        // Imposta l'intervallo per i tentativi successivi (ogni ora)
        setInterval(autoImportPP, 60 * 60 * 1000);
    }




// VERSIONE AGGIORNATA (DOPO LA MODIFICA)
function showSplitTooltip(taskData, event) {
    if (!genericTooltip) genericTooltip = document.getElementById('genericTooltip');
    if (!taskData) return;

    // Aggiungiamo le unità di misura e le informazioni aggiuntive (OP/OV/Lotto)
    const productionInfoParts = [];
    // Informazioni generali
    productionInfoParts.push(`<strong>Codice:</strong> ${taskData.codice || 'N/D'}`);
    if (taskData.op) productionInfoParts.push(`<strong>OP:</strong> ${taskData.op}`);
    if (taskData.ov) productionInfoParts.push(`<strong>OV:</strong> ${taskData.ov}`);
    if (taskData.lotto) productionInfoParts.push(`<strong>Lotto:</strong> ${taskData.lotto}`);
    productionInfoParts.push(`<strong>Prodotto:</strong> ${taskData.prodotto || 'N/D'}`);
    productionInfoParts.push(`<strong>Cliente:</strong> ${taskData.cliente || 'N/D'}`);
    productionInfoParts.push(`<strong>Qtà Richiesta:</strong> ${taskData.quantitaRichiesta ? taskData.quantitaRichiesta + ' ' + taskData.quantitaRichiestaUnit : 'N/D'}`);
    productionInfoParts.push(`<strong>Giacenza:</strong> ${taskData.giacenzaMagazzino ? taskData.giacenzaMagazzino + ' Kg' : 'N/D'}`);
    productionInfoParts.push(`<strong>Da Produrre:</strong> ${taskData.quantitaDaProdurre ? taskData.quantitaDaProdurre + ' Kg' : 'N/D'}`);
    productionInfoParts.push(`<strong>Materie Prime:</strong> ${taskData.materiePrime || 'N/D'}`);
    productionInfoParts.push(`<strong>Macchinario:</strong> ${taskData.macchinari || 'N/D'}`);
    productionInfoParts.push(`<strong>Operatore:</strong> ${taskData.operatore || 'N/D'}`);
    productionInfoParts.push(`<strong>Data Prod.:</strong> ${taskData.produzioneData || 'N/D'}`);
    productionInfoParts.push(`<strong>Giorni Prod.:</strong> ${taskData.giorniDiProduzione || 'N/D'}`);
    const productionInfo = productionInfoParts.join('<br>');

    let packagingInfo = [
        `<strong>Pezzi:</strong> ${taskData.confezionamentoPezzi || 'N/D'}`,
        `<strong>Kg/Pezzo:</strong> ${taskData.confezionamentoKgPerPiece || 'N/D'}`,
        `<strong>Data Confez.:</strong> ${taskData.dataConfezionamento || 'N/D'}`,
        `<strong>Cod. Confez.:</strong> ${taskData.codiceConfezionamento || 'N/D'}`,
        `<strong>Lotto SC:</strong> ${taskData.lottoSC || 'N/D'}`,
        `<strong>Materiale Confez.:</strong> ${taskData.materialeConfezionamento || 'N/D'}`,
        `<strong>Data Sped.:</strong> ${taskData.dataSpedizione || 'N/D'}`,
        `<strong>Note:</strong> ${taskData.note || 'N/D'}`
    ].join('<br>');

    // Aggiunge eventuali dettagli DeviceRef se esiste una corrispondenza per il codice della produzione.
    try {
        const deviceRefs = typeof getDeviceRefData === 'function' ? getDeviceRefData() : JSON.parse(localStorage.getItem('deviceRefData') || '[]');
        const codeKey = String(taskData.codice || '').trim().toUpperCase();
        const matchedRef = Array.isArray(deviceRefs) ? deviceRefs.find(ref => String(ref.codice || '').trim().toUpperCase() === codeKey) : null;
        if (matchedRef) {
            const lines = [];
            if (matchedRef.aghiPresenti) lines.push(`<strong>Aghi presenti:</strong> ${matchedRef.aghiPresenti}`);
            if (matchedRef.tipologiaAghi) lines.push(`<strong>Tipologia aghi:</strong> ${matchedRef.tipologiaAghi}`);
            // Etichette modificate: rimuoviamo il cancelletto per "aghi per valva" e "siringhe/scatola"
            if (matchedRef.aghiPerValva) lines.push(`<strong>Aghi per valva:</strong> ${matchedRef.aghiPerValva}`);
            if (matchedRef.siringhePerScatola) lines.push(`<strong>Siringhe/scatola:</strong> ${matchedRef.siringhePerScatola}`);
            if (matchedRef.volumeMl) lines.push(`<strong>Volume siringa:</strong> ${matchedRef.volumeMl} mL`);
            // Mostra il numero di scatole/scatolone (colonna I) con etichetta più intuitiva.
            // In ordine di priorità: usa il nuovo campo pezziPerScatolone, altrimenti siringhePerScatola2 o siringhePerScatola.
            const scatolePerScatolone = (matchedRef.pezziPerScatolone !== undefined ? matchedRef.pezziPerScatolone : undefined) ||
                                        matchedRef.siringhePerScatola2 || matchedRef.siringhePerScatola;
            if (scatolePerScatolone) {
                // Etichetta aggiornata: mostra il numero di scatole/scatolone (colonna I)
                lines.push(`<strong>N° scatole/scatolone:</strong> ${scatolePerScatolone}`);
            }
            // Mostra i pesi singoli, se disponibili
            if (matchedRef.pesoScatola) lines.push(`<strong>Peso scatola:</strong> ${matchedRef.pesoScatola}`);
            if (matchedRef.pesoScatolone) lines.push(`<strong>Peso scatolone:</strong> ${matchedRef.pesoScatolone}`);

            // Calcolo del numero teorico di scatoloni e del peso totale ipotetico.
            try {
                const totalPieces = parseFloat(taskData.confezionamentoPezzi || taskData.quantita || 0);
                // Determina il numero di pezzi per scatolone: utilizza prima il campo pezziPerScatolone (colonna I),
                // altrimenti ricade su siringhePerScatola2 o, in ultima istanza, su siringhePerScatola.
                const pezziPerScatolone = parseFloat(
                    (matchedRef.pezziPerScatolone !== undefined ? matchedRef.pezziPerScatolone : undefined) ||
                    matchedRef.siringhePerScatola2 || matchedRef.siringhePerScatola || 0
                );
                if (!isNaN(totalPieces) && !isNaN(pezziPerScatolone) && pezziPerScatolone > 0) {
                    const theoreticalBoxes = Math.ceil(totalPieces / pezziPerScatolone);
                    lines.push(`<strong># scatoloni (ip.):</strong> ${theoreticalBoxes}`);
                    // Calcolo del peso totale ipotetico: moltiplica il numero teorico di scatoloni per il peso di ogni scatolone.
                    const pesoPerScatolone = parseFloat(matchedRef.pesoScatolone || matchedRef.pesoScatola || 0);
                    if (!isNaN(pesoPerScatolone) && pesoPerScatolone > 0) {
                        const totalWeight = (pesoPerScatolone * theoreticalBoxes).toFixed(2);
                        lines.push(`<strong>Peso totale (ip.):</strong> ${totalWeight}`);
                    }
                }
            } catch (calcErr) {
                console.warn('Errore nel calcolo del numero teorico di scatoloni o del peso:', calcErr);
            }
            // Aggiunge sempre i campi Quantità reale e Numero scatoloni ipotetici
            try {
                let mdQtyVal = '';
                let mdBoxesVal = '';
                const mdData = typeof getMedicalProductionData === 'function'
                    ? getMedicalProductionData()
                    : JSON.parse(localStorage.getItem('medicalProductionData') || '[]');
                const codeKeyMD = codeKey;
                // Lotto per MD, se disponibile in taskData
                const lotKeyMD = String(taskData.lotto || '').trim().toUpperCase();
                const mdMatches = Array.isArray(mdData)
                    ? mdData.filter(mp => {
                        const sameCode = String(mp.codice || '').trim().toUpperCase() === codeKeyMD;
                        if (lotKeyMD) {
                            return sameCode && String(mp.lotto || '').trim().toUpperCase() === lotKeyMD;
                        }
                        return sameCode;
                    })
                    : [];
                if (mdMatches.length > 0) {
                    // Se possibile seleziona la riga MD con la stessa data di produzione dell'OPI corrente
                    let mdItem = null;
                    try {
                        const opiProdDateKey = String(taskData.dataProd || '').trim();
                        if (opiProdDateKey) {
                            mdItem = mdMatches.find(mp => {
                                const mdDate = String(mp.data || mp.produzioneData || '').trim();
                                return mdDate === opiProdDateKey;
                            }) || null;
                        }
                    } catch (selErr) {
                        mdItem = null;
                    }
                    if (!mdItem) {
                        mdItem = mdMatches[0];
                    }
                    // Preleva il valore "Quantità (pz)" direttamente dalla colonna "quantita" della tabella MD.
                    let qtyPieces = 0;
                    if (mdItem && mdItem.quantita !== undefined && mdItem.quantita !== null && mdItem.quantita !== '') {
                        const parsedQty = parseFloat(String(mdItem.quantita).replace(/\./g, '').replace(',', '.'));
                        if (!isNaN(parsedQty) && parsedQty > 0) {
                            qtyPieces = parsedQty;
                        }
                    }
                    if (!isNaN(qtyPieces) && qtyPieces > 0) mdQtyVal = qtyPieces;
                    // Numero di pezzi per scatolone dal DeviceRef (campo pezziPerScatolone o fallback)
                    let perBoxValCalc;
                    const rawPerBoxMD = (matchedRef.pezziPerScatolone !== undefined ? matchedRef.pezziPerScatolone : undefined) ||
                                         matchedRef.siringhePerScatola2 || matchedRef.siringhePerScatola;
                    if (rawPerBoxMD != null && rawPerBoxMD !== '') {
                        const normalizedPB = String(rawPerBoxMD).replace(/\./g, '').replace(',', '.');
                        const numPB = parseFloat(normalizedPB);
                        if (!isNaN(numPB) && numPB > 0) perBoxValCalc = numPB;
                    }
                    if (perBoxValCalc && !isNaN(qtyPieces) && qtyPieces > 0) {
                        mdBoxesVal = Math.ceil(qtyPieces / perBoxValCalc);
                    }
                }
                lines.push(`<strong>Quantità reale siringhe:</strong> ${mdQtyVal || 'N/D'}`);
                lines.push(`<strong>Numero scatoloni ipotetici:</strong> ${mdBoxesVal || 'N/D'}`);
            } catch (calcErr2) {
                console.warn('Errore nel calcolo MD per tooltip produzione:', calcErr2);
                lines.push(`<strong>Quantità reale siringhe:</strong> N/D`);
                lines.push(`<strong>Numero scatoloni ipotetici:</strong> N/D`);
            }
            if (lines.length > 0) {
                packagingInfo += '<br><br><strong>Info Dispositivo:</strong><br>' + lines.join('<br>');
            }
            // Recupera eventuali informazioni dal Programma Giornaliero di Produzione e le aggiunge al tooltip.
            try {
                let dailyInfoStr = '';
                const dailyRows = document.querySelectorAll('#dailyProductionTable tbody tr');
                const codeKeyDaily = codeKey;
                const ovKeyDaily = String(taskData.ov || '').trim().toUpperCase();
                const opKeyDaily = String(taskData.op || '').trim().toUpperCase();
                const matchingDaily = [];
                dailyRows.forEach(dr => {
                    const d = getDailyRowData(dr);
                    const dOv = String(d.ov || '').trim().toUpperCase();
                    const dOp = String(d.op || '').trim().toUpperCase();
                    const dCode = String(d.codice || '').trim().toUpperCase();
                    // Coincide per codice articolo e OV (se presente) oppure per codice e OP (se presente)
                    if (dCode === codeKeyDaily && ((ovKeyDaily && dOv === ovKeyDaily) || (opKeyDaily && dOp === opKeyDaily))) {
                        matchingDaily.push(d);
                    }
                });
                if (matchingDaily.length > 0) {
                    dailyInfoStr = matchingDaily.map(d => {
                        const parts = [];
                        if (d.ov) parts.push(`<strong>OV:</strong> ${d.ov}`);
                        if (d.op) parts.push(`<strong>OP:</strong> ${d.op}`);
                        if (d.lotto) parts.push(`<strong>Lotto:</strong> ${d.lotto}`);
                        if (d.quantita) parts.push(`<strong>Quantità:</strong> ${d.quantita}`);
                        if (d.quantitaConfezionamento) parts.push(`<strong>Quantità confezionamento:</strong> ${d.quantitaConfezionamento}`);
                        if (d.macchinario) parts.push(`<strong>Macchinario:</strong> ${d.macchinario}`);
                        if (d.operazioni) parts.push(`<strong>Operazioni:</strong> ${d.operazioni}`);
                        if (d.operatore) parts.push(`<strong>Operatore:</strong> ${d.operatore}`);
                        return parts.join('<br>');
                    }).join('<br><br>');
                }
                if (dailyInfoStr) {
                    packagingInfo += '<br><br><strong>Programma Giornaliero:</strong><br>' + dailyInfoStr;
                }
            } catch (dailyErr) {
                console.warn('Errore nel recupero dati dal Programma Giornaliero per tooltip produzione:', dailyErr);
            }
        }
    } catch (e) {
        console.warn('Errore nel recupero DeviceRef per tooltip produzione:', e);
    }
    
    genericTooltip.innerHTML = `
        <div class="tooltip-container">
            <div class="tooltip-box production-tooltip"><h3>Dettaglio Produzione</h3><p>${productionInfo}</p></div>
            <div class="tooltip-box packaging-tooltip"><h3>Dettaglio Confezionamento</h3><p>${packagingInfo}</p></div>
        </div>`;
    
    genericTooltip.style.backgroundColor = 'transparent';
    genericTooltip.style.color = '#333';
    genericTooltip.style.padding = '0';
    genericTooltip.style.border = 'none';

    genericTooltip.classList.add('visible');
    moveGenericTooltip(event);
}

// Funzione DEFINITIVA E COMPLETA per mostrare il tooltip delle spedizioni
function showSplitShippingTooltip(taskData, event) {
    if (!genericTooltip) genericTooltip = document.getElementById('genericTooltip');
    if (!taskData) return;

    // Imposta lo stile del contenitore del tooltip
    genericTooltip.style.backgroundColor = 'transparent';
    genericTooltip.style.padding = '0';
    genericTooltip.style.border = 'none';
    genericTooltip.style.boxShadow = '0 4px 15px rgba(0,0,0,0.25)';
    genericTooltip.style.maxWidth = '950px'; // Larghezza sufficiente per 3 box

    // Prepara il contenuto per il Box 1: Dettaglio Ordine
    // Determina l'etichetta corretta per l'ordine: per le spedizioni (con rowId) mostra "OV", per gli arrivi (senza rowId) mostra "OA".
    const ovLabel = taskData && taskData.rowId ? 'OV' : 'OA';
    const orderInfoArray = [
        `<strong>${ovLabel}:</strong> ${taskData.ov || 'N/D'}`,
        `<strong>Codice Articolo:</strong> ${taskData.codiceArticolo || 'N/D'}`,
        `<strong>Descrizione:</strong> ${taskData.descrizioneArticolo || 'N/D'}`,
        `<strong>Quantità:</strong> ${taskData.quantita || 'N/D'} ${taskData.um || ''}`,
        `<strong>Data Consegna:</strong> ${taskData.dataConsegna || 'N/D'}`
    ];

    // Solo per la merce in arrivo (task senza rowId) aggiungi il layout di riferimento.
    // Se il layout contiene G5CELLA e una temperatura (+4°C, -20°C o +20°C) evidenzia in rosso,
    // altrimenti visualizza in nero per distinguere rapidamente le celle a temperatura controllata.
    if (!taskData.rowId) {
        const layoutStr = String(taskData.layout || '').trim();
        if (layoutStr) {
            const upperLayout = layoutStr.toUpperCase();
            let layoutColor = 'black';
            if (
                (upperLayout.includes('G5CELLA +4°C')) ||
                (upperLayout.includes('G5CELLA -20°C')) ||
                (upperLayout.includes('G5CELLA +20°C'))
            ) {
                layoutColor = 'red';
            }
            orderInfoArray.push(`<strong>Layout:</strong> <span style="color:${layoutColor};">${layoutStr}</span>`);
        } else {
            orderInfoArray.push(`<strong>Layout:</strong> N/D`);
        }
    }

    const orderInfo = orderInfoArray.join('<br>');

    // Prepara il contenuto per il Box 2: Dettaglio Cliente
    const clientInfo = [
        `<strong>Ragione Sociale:</strong> ${taskData.ragioneSociale || 'N/D'}`,
        `<strong>Rif. Cliente:</strong> ${taskData.riferimentoCliente || 'N/D'}`,
        `<strong>Indirizzo:</strong> ${taskData.indirizzo || 'N/D'}`,
        `<strong>Città:</strong> ${taskData.citta || 'N/D'} (${taskData.provincia || 'N/D'})`
    ].join('<br>');

    // Prepara il contenuto per il Box dei commenti solo se esiste un rowId (spedizioni)
    let commentBoxHtml = '';
    if (taskData.rowId) {
        const commentiHtml = taskData.commentiQA ?
            `<p style="white-space: pre-wrap;">${taskData.commentiQA}</p>` :
            '<p>Nessun commento.</p>';
        commentBoxHtml = `
            <div class="tooltip-box qa-comments-tooltip">
                <h3>Commenti QA</h3>
                ${commentiHtml}
            </div>`;
    }

    // Prepara il contenuto per il Box OPI e gli stati CQ/QA solo per le spedizioni (task con rowId)
    let opiBoxHtml = '';
    if (taskData.rowId) {
        try {
            const opiData = typeof getOpiMonitorData === 'function' ? getOpiMonitorData() : JSON.parse(localStorage.getItem('opi_monitor_data') || '[]');
            // Filtra gli ordini di produzione (OPI) per il tooltip delle spedizioni.
            // Vengono selezionati solo i record il cui codice articolo compare
            // nel programma di spedizione per l'ordine di vendita (OV) corrente.
            // Non si utilizzano più i dati di produzione per incrociare le date.
            const shippingRows = (typeof getAllShippingData === 'function') ? getAllShippingData() : [];
            /*
             * Filtra le righe di spedizione per l'OV corrente e per la data
             * di consegna del task corrente.  Deduplica i codici articolo
             * in modo da evitare ripetizioni quando un articolo appare in
             * più righe nella stessa data.  Memorizza anche la data di
             * consegna per ogni codice, da utilizzare per l'abbinamento
             * successivo con la scadenza dell'OPI.
             */
            let codesForOv = [];
            const seenCodes = new Set();
            shippingRows.forEach(row => {
                const ovMatch = String(row.ov || '').trim().toUpperCase() === String(taskData.ov || '').trim().toUpperCase();
                const dateMatch = String(row.dataConsegna || '').trim() === String(taskData.dataConsegna || '').trim();
                if (!ovMatch || !dateMatch) return;
                const codeKey = String(row.codiceArticolo || '').trim().toUpperCase();
                if (seenCodes.has(codeKey)) return;
                seenCodes.add(codeKey);
                codesForOv.push({
                    code: codeKey,
                    consegna: String(row.dataConsegna || '').trim()
                });
            });

            // Se i dati della packing list sono disponibili, limita ulteriormente i codici
            // ai soli presenti nella packing list per l'OV corrente e per la data di consegna del task.
            if (typeof window !== 'undefined' && window.packingListData) {
                try {
                    const listForOv = window.packingListData[String(taskData.ov || '').trim()] || [];
                    if (Array.isArray(listForOv) && listForOv.length > 0) {
                        const validCodes = new Set();
                        listForOv.forEach(item => {
                            const itemCode = String(item.codiceArticolo || '').trim().toUpperCase();
                            if (itemCode) validCodes.add(itemCode);
                        });
                        if (validCodes.size > 0) {
                            codesForOv = codesForOv.filter(obj => validCodes.has(obj.code));
                        }
                    }
                } catch (_) {}
            }
            /*
             * Costruisci l'elenco degli ordini di produzione da mostrare nel tooltip.
             * Per ogni articolo presente nel programma di spedizione (codesForOv)
             * viene selezionato il primo OPI associato all'ordine di vendita e al
             * codice articolo corrispondente. In questo modo si evita di
             * visualizzare gli altri ordini di produzione non pertinenti alla
             * spedizione della data corrente, mantenendo al contempo tutte le
             * informazioni (OV, OP, lotto, date, numero pezzi, etc.) per il
             * singolo OPI scelto.  Questa logica ricalca il comportamento della
             * packing list, dove viene utilizzato un solo OPI per ciascun
             * articolo nelle spedizioni.
             */
            const matchedOpi = [];
            codesForOv.forEach(item => {
                // Seleziona il primo OPI associato all'OV corrente e al
                // codice articolo corrente, senza imporre una corrispondenza
                // con la data di scadenza.  Questa logica ricalca il
                // comportamento della packing list, che mostra un solo
                // ordine di produzione per ciascun articolo collegato
                // all'ordine di vendita.
                const match = opiData.find(opi => {
                    return String(opi.ov || '').trim().toUpperCase() === String(taskData.ov || '').trim().toUpperCase() &&
                           String(opi.codice || '').trim().toUpperCase() === item.code;
                }) || null;
                if (match) matchedOpi.push(match);
            });
            // Mappa degli stati CQ e QA per trasformare il codice colore in un testo descrittivo
            const statusCQMap = {
                white: 'Merce da analizzare / in analisi',
                yellow: 'Merce accettata con deroga',
                green: 'Merce conforme',
                red: 'Merce non conforme'
            };
            const statusQAMap = {
                white: 'Merce in fase di valutazione',
                yellow: 'Merce accettata con deroga',
                green: 'Merce conforme',
                red: 'Merce non conforme'
            };
            // Mappa colori corrispondente agli stati CQ/QA.  In base allo stato
            // (white, yellow, green, red) viene selezionato un colore diverso per
            // il testo visualizzato nel tooltip.  Il colore neutro per lo stato
            // "white" è scuro per garantire leggibilità su sfondo chiaro.
            const statusColorMap = {
                white: '#333',    // nero/grigio scuro per "Merce da analizzare"
                yellow: '#FFC107',// giallo intenso per "accettata con deroga"
                green: '#4CAF50', // verde per "conforme"
                red: '#F44336'    // rosso per "non conforme"
            };
            const cqText = statusCQMap[taskData.cqStatus || 'white'] || '';
            const qaText = statusQAMap[taskData.qaStatus || 'white'] || '';
            // Prepara l'elenco degli stati CQ e QA con colori dinamici.  Ogni
            // stato utilizza il colore definito nella mappa statusColorMap,
            // consentendo al testo di riflettere visivamente il colore del
            // pallino/flag corrispondente.
            const cqColor = statusColorMap[taskData.cqStatus || 'white'] || '#333';
            const qaColor = statusColorMap[taskData.qaStatus || 'white'] || '#333';
            // Se il codice articolo rientra nell'elenco ADR, visualizza una
            // avvertenza in rosso prima degli stati.  L'avviso viene
            // visualizzato nel tooltip solo per le spedizioni.
            // Verifica se il codice è ADR leggendo la lista globale.  Se la lista
            // non è presente (ad es. non ancora caricata), considera false.
            const isAdrCode = (window.adrCodes && window.adrCodes.has(String(taskData.codiceArticolo || '').trim().toUpperCase()));
            const adrWarningLine = isAdrCode
                ? `<strong><span style="color:#F44336; text-decoration: underline;">ATTENZIONE: Trasporto ADR</span></strong>`
                : '';
            const statusLines = [];
            if (adrWarningLine) statusLines.push(adrWarningLine);
            statusLines.push(
                `<strong>Stato CQ:</strong> <span style="color:${cqColor};">${cqText}</span>`,
                `<strong>Stato QA:</strong> <span style="color:${qaColor};">${qaText}</span>`
            );
            let opiDetails = '';
            // Recupera eventuali informazioni DeviceRef per il codice articolo corrente
            let deviceRefDetails = '';
            try {
                const deviceRefs = typeof getDeviceRefData === 'function' ? getDeviceRefData() : JSON.parse(localStorage.getItem('deviceRefData') || '[]');
                const codeKey = String(taskData.codiceArticolo || '').trim().toUpperCase();
                const matchedRef = Array.isArray(deviceRefs) ? deviceRefs.find(ref => String(ref.codice || '').trim().toUpperCase() === codeKey) : null;
                if (matchedRef) {
                    const fields = [];
                    // Aggiunge le informazioni solo se presenti.  Le etichette sono coerenti
                    // con l'ordine delle colonne importate: cliente, aghi presenti, tipologia aghi,
                    // aghi per valva, siringhe per scatola, volume (ml), siringhe per scatola 2,
                    // peso scatola, peso scatolone.
                    if (matchedRef.cliente) fields.push(`<strong>Cliente (Ref):</strong> <span style="color:#333;">${matchedRef.cliente}</span>`);
                    if (matchedRef.aghiPresenti) fields.push(`<strong>Aghi presenti:</strong> <span style="color:#333;">${matchedRef.aghiPresenti}</span>`);
                    if (matchedRef.tipologiaAghi) fields.push(`<strong>Tipologia aghi:</strong> <span style="color:#333;">${matchedRef.tipologiaAghi}</span>`);
                    // Etichetta modificata: rimuove il cancelletto per "aghi per valva"
                    if (matchedRef.aghiPerValva) fields.push(`<strong>Aghi per valva:</strong> <span style="color:#333;">${matchedRef.aghiPerValva}</span>`);
                    // Etichetta modificata: rimuove il cancelletto per "siringhe/scatola"
                    if (matchedRef.siringhePerScatola) fields.push(`<strong>Siringhe/scatola:</strong> <span style="color:#333;">${matchedRef.siringhePerScatola}</span>`);
                    // Aggiungi il volume con unità "mL" se disponibile
                    if (matchedRef.volumeMl) fields.push(`<strong>Volume siringa:</strong> <span style="color:#333;">${matchedRef.volumeMl} mL</span>`);

                    // Calcola i valori di produzione medicale (quantità e scatoloni ipotetici) se disponibili
let mdQtyVal = '';
let mdBoxesVal = '';
try {
    const mdData = typeof getMedicalProductionData === 'function'
        ? getMedicalProductionData()
        : JSON.parse(localStorage.getItem('medicalProductionData') || '[]');

    const codeKeyMD = codeKey;
    const isMdLike = codeKeyMD.replace('*','').startsWith('4') || ['7545','40125V','7316','7317'].includes(codeKeyMD);

    if (isMdLike) {
        const mdMatches = Array.isArray(mdData)
            ? mdData.filter(mp => String(mp.codice || '').trim().toUpperCase() === codeKeyMD)
            : [];

        let selectedMdItem = null;

        // 1) Prova abbinamento per lotto dall'OPI (o lotto bulk derivato)
        let opiLotKey = '';
        if (Array.isArray(matchedOpi) && matchedOpi.length > 0) {
            opiLotKey = String(matchedOpi[0].lotto || '').trim().toUpperCase();
        }
        const bulkLotKey = opiLotKey ? deriveBulkLotto(opiLotKey) : '';

        if (mdMatches.length > 0) {
            if (opiLotKey || bulkLotKey) {
                selectedMdItem = mdMatches.find(mp => {
                    const l = String(mp.lotto || '').trim().toUpperCase();
                    return (opiLotKey && l.includes(opiLotKey)) || (bulkLotKey && l.includes(bulkLotKey));
                }) || null;
            }
            // 2) Se non c'è match per lotto, prova con la data di produzione dell'OPI
            if (!selectedMdItem) {
                let opiProdDateKey = '';
                if (Array.isArray(matchedOpi) && matchedOpi.length > 0) {
                    opiProdDateKey = String(matchedOpi[0].dataProd || '').trim();
                }
                if (opiProdDateKey) {
                    selectedMdItem = mdMatches.find(mp => String(mp.data || '').trim() === opiProdDateKey) || null;
                }
            }
            // 3) Fallback: prendi l'ultima riga MD (per data) per quel codice
            if (!selectedMdItem) {
                selectedMdItem = mdMatches.slice().sort((a,b) => {
                    const da = (typeof flatpickr !== 'undefined' && flatpickr.parseDate) ? (flatpickr.parseDate(a.data, "d/m/Y") || new Date(0)) : new Date(0);
                    const db = (typeof flatpickr !== 'undefined' && flatpickr.parseDate) ? (flatpickr.parseDate(b.data, "d/m/Y") || new Date(0)) : new Date(0);
                    return db - da;
                })[0] || null;
            }
        }

        // Quantità reale (pezzi) direttamente da "quantita" della riga MD selezionata
        let qtyPieces = 0;
        if (selectedMdItem && selectedMdItem.quantita !== undefined && selectedMdItem.quantita !== null && selectedMdItem.quantita !== '') {
            const parsedQ = parseFloat(String(selectedMdItem.quantita).replace(/\./g, '').replace(',', '.'));
            if (!isNaN(parsedQ) && parsedQ > 0) qtyPieces = parsedQ;
        }

        // 4) Fallback dalla packing list (se disponibile in memoria)
        if (!(qtyPieces > 0) && typeof window !== 'undefined' && window.packingListData) {
            try {
                const ovKey = String(taskData.ov || '').trim();
                const list = window.packingListData[ovKey] || [];
                const itemPL = list.find(it => String(it.codiceArticolo || '').trim().toUpperCase() === codeKeyMD);
                const val = itemPL && (itemPL.quantitaReale || itemPL.pezziReali);
                if (val) {
                    const parsed = parseFloat(String(val).replace(/\./g, '').replace(',', '.'));
                    if (!isNaN(parsed) && parsed > 0) qtyPieces = parsed;
                }
            } catch (_) {}
        }

        if (qtyPieces > 0) {
            mdQtyVal = qtyPieces;
            // Numero scatoloni ipotetici in base al valore "per scatola/scatolone" più affidabile
            let perBoxValCalc;
            const rawPerBoxMD = (matchedRef.pezziPerScatolone !== undefined ? matchedRef.pezziPerScatolone : undefined) ||
                                matchedRef.siringhePerScatola2 || matchedRef.siringhePerScatola;
            if (rawPerBoxMD != null && rawPerBoxMD !== '') {
                const normalizedPB = String(rawPerBoxMD).replace(/\./g, '').replace(',', '.');
                const numPB = parseFloat(normalizedPB);
                if (!isNaN(numPB) && numPB > 0) perBoxValCalc = numPB;
            }
            if (perBoxValCalc) {
                mdBoxesVal = Math.ceil(qtyPieces / perBoxValCalc);
            }
        }
    }
} catch (calcErr) {
    console.warn('Errore nel calcolo dei dati MD per tooltip spedizioni:', calcErr);
}
                    // Mostra il numero di siringhe per scatolone (colonna I di DeviceRef) con etichetta corretta.
                    // In ordine di priorità utilizziamo il campo pezziPerScatolone, altrimenti siringhePerScatola2 o siringhePerScatola.
                    {
                        let siringhePerScatoloneVal = '';
                        const rawVal1 = matchedRef.pezziPerScatolone !== undefined ? matchedRef.pezziPerScatolone : undefined;
                        const rawVal2 = matchedRef.siringhePerScatola2 || matchedRef.siringhePerScatola;
                        if (rawVal1 != null && rawVal1 !== '') {
                            siringhePerScatoloneVal = rawVal1;
                        } else if (rawVal2 != null && rawVal2 !== '') {
                            siringhePerScatoloneVal = rawVal2;
                        }
                        if (siringhePerScatoloneVal) {
                            // Etichetta aggiornata: rinomina il campo per evitare confusione con il numero di siringhe; indica il numero di scatole/scatolone
                            fields.push(`<strong>N° scatole/scatolone:</strong> <span style="color:#333;">${siringhePerScatoloneVal}</span>`);
                        }
                    }
                    if (matchedRef.pesoScatola) fields.push(`<strong>Peso scatola:</strong> <span style="color:#333;">${matchedRef.pesoScatola}</span>`);
                    if (matchedRef.pesoScatolone) fields.push(`<strong>Peso scatolone:</strong> <span style="color:#333;">${matchedRef.pesoScatolone}</span>`);
                    // Aggiunge sempre i campi Quantità reale e Numero scatoloni ipotetici
                    fields.push(`<strong>Quantità reale siringhe:</strong> <span style="color:#333;">${mdQtyVal || 'N/D'}</span>`);
                    fields.push(`<strong>Numero scatoloni ipotetici:</strong> <span style="color:#333;">${mdBoxesVal || 'N/D'}</span>`);
                    if (fields.length > 0) {
                        deviceRefDetails = fields.join('<br>');
                    }
                }
            } catch (e) {
                console.warn('Errore nel recupero dei dati DeviceRef:', e);
            }
            if (matchedOpi.length > 0) {
                /*
                 * Deduplica le righe OPI identiche prima di formattarle e
                 * rimuove il campo Operatore, non necessario in fase di
                 * spedizione.  Vengono considerate uguali le righe con gli
                 * stessi valori di OV, OP, lotto, data di produzione, scadenza,
                 * quantità e unità di misura.
                 */
                const uniqueOpi = [];
                const seenOpiKeys = new Set();
                matchedOpi.forEach(item => {
                    const key = [
                        String(item.ov || ''),
                        String(item.op || ''),
                        String(item.lotto || ''),
                        String(item.dataProd || ''),
                        String(item.scadenza || ''),
                        String(item.quantita || ''),
                        String(item.um || '')
                    ].join('|');
                    if (!seenOpiKeys.has(key)) {
                        seenOpiKeys.add(key);
                        uniqueOpi.push(item);
                    }
                });
                // Raggruppa le righe OPI per codice articolo e descrizione (articolo) in modo
                // da aggiungere un'intestazione sopra ciascun gruppo.  Questo rende
                // più immediata la comprensione di quale lotto/OP appartiene a
                // quale articolo, specialmente quando un OV ha più articoli diversi.
                const opiByArticle = {};
                uniqueOpi.forEach(item => {
                    const codeKey = String(item.codice || '').trim().toUpperCase();
                    const descrKey = String(item.articolo || item.descrizione || '').trim();
                    const groupKey = `${codeKey}|${descrKey}`;
                    if (!opiByArticle[groupKey]) opiByArticle[groupKey] = [];
                    opiByArticle[groupKey].push(item);
                });
                opiDetails = Object.keys(opiByArticle).map(groupKey => {
                    const [codeKey, descrKey] = groupKey.split('|');
                    const items = opiByArticle[groupKey];
                    // Intestazione con codice articolo e descrizione
                    let details = `<span style="font-weight:bold;color:#333;">${codeKey}`;
                    if (descrKey) {
                        details += ` - ${descrKey}`;
                    }
                    details += `</span>`;
                    // Aggiungi le righe OPI per il gruppo
                    details += '<br>' + items.map(item => {
                        const opVal = item.op || 'N/D';
                        const lottoVal = item.lotto || 'N/D';
                        const qtyVal = (item.quantita !== undefined && item.quantita !== null && item.quantita !== '') ? item.quantita : 'N/D';
                        const umVal = (item.um !== undefined && item.um !== null && item.um !== '') ? item.um : '';
                        const dataProdVal = item.dataProd || 'N/D';
                        const scadenzaVal = item.scadenza || 'N/D';
                        return `<strong>OP:</strong> <span style="color:#333;">${opVal}</span><br>` +
                               `<strong>Lotto:</strong> <span style="color:#333;">${lottoVal}</span><br>` +
                               `<strong>Numero pezzi:</strong> <span style="color:#333;">${qtyVal} ${umVal}</span><br>` +
                               `<strong>Data Prod.:</strong> <span style="color:#333;">${dataProdVal}</span><br>` +
                               `<strong>Scadenza:</strong> <span style="color:#333;">${scadenzaVal}</span>`;
                    }).join('<br><br>');
                    return details;
                }).join('<br><br>');
            }
            // Recupera eventuali informazioni dal Programma Giornaliero di Produzione (daily production)
            let dailyDetails = '';
            try {
                const dailyRows = document.querySelectorAll('#dailyProductionTable tbody tr');
                const codeKeyDaily = String(taskData.codiceArticolo || '').trim().toUpperCase();
                const ovKeyDaily = String(taskData.ov || '').trim().toUpperCase();
                const matchingDaily = [];
                dailyRows.forEach(dr => {
                    const d = getDailyRowData(dr);
                    const dOv = String(d.ov || '').trim().toUpperCase();
                    const dCode = String(d.codice || '').trim().toUpperCase();
                    if (dOv === ovKeyDaily && dCode === codeKeyDaily) {
                        matchingDaily.push(d);
                    }
                });
                if (matchingDaily.length > 0) {
                    dailyDetails = matchingDaily.map(d => {
                        const ovVal = d.ov || 'N/D';
                        const opVal = d.op || 'N/D';
                        const lottoVal = d.lotto || 'N/D';
                        const operatoreVal = d.operatore || 'N/D';
                        return `<strong>OV:</strong> <span style="color:#333;">${ovVal}</span><br>` +
                               `<strong>OP:</strong> <span style="color:#333;">${opVal}</span><br>` +
                               `<strong>Lotto:</strong> <span style="color:#333;">${lottoVal}</span>`;
                    }).join('<br><br>');
                }
            } catch (dailyErr) {
                console.warn('Errore nel recupero dati dal Programma Giornaliero:', dailyErr);
            }
            // Combina le informazioni OPI, DeviceRef e Daily (programma giornaliero) nell'ordine corretto
            if (deviceRefDetails) {
                if (opiDetails) {
                    opiDetails = `${opiDetails}<br><br>${deviceRefDetails}`;
                } else {
                    opiDetails = deviceRefDetails;
                }
            }
            if (dailyDetails) {
                if (opiDetails) {
                    opiDetails = `${opiDetails}<br><br>${dailyDetails}`;
                } else {
                    opiDetails = dailyDetails;
                }
            }

            // Recupera eventuali informazioni di produzione medicale (MedicalProductionData) per il codice corrente
            let mdDetails = '';
            try {
                const mdData = typeof getMedicalProductionData === 'function'
                    ? getMedicalProductionData()
                    : JSON.parse(localStorage.getItem('medicalProductionData') || '[]');
                // Calcola la chiave MD basandosi sul codice articolo del task corrente
                const codeKeyMD = String(taskData.codiceArticolo || '').trim().toUpperCase();
                const mdMatches = Array.isArray(mdData)
                    ? mdData.filter(mp => String(mp.codice || '').trim().toUpperCase() === codeKeyMD)
                    : [];
                if (mdMatches.length > 0) {
                    mdDetails = mdMatches.map(mp => {
                        const parts = [];
                        if (mp.data) parts.push(`<strong>Data prod.:</strong> <span style="color:#333;">${mp.data}</span>`);
                        if (mp.lotto) parts.push(`<strong>Lotto:</strong> <span style="color:#333;">${mp.lotto}</span>`);
                        parts.push(`<strong>Quantità:</strong> <span style="color:#333;">${mp.quantita}</span>`);
                        parts.push(`<strong>Pezzi reali:</strong> <span style="color:#333;">${mp.unita}</span>`);
                        return parts.join('<br>');
                    }).join('<br><br>');
                }
            } catch (e) {
                console.warn('Errore nel recupero dei dati Medical Production:', e);
            }
            // Se esistono dettagli produzione MD, integrali dopo i dettagli OPI/DeviceRef
            if (mdDetails) {
                if (opiDetails) {
                    opiDetails = `${opiDetails}<br><br><strong>Produzione MD:</strong><br>${mdDetails}`;
                } else {
                    opiDetails = `<strong>Produzione MD:</strong><br>${mdDetails}`;
                }
            }

            // Combina i dettagli OPI/DeviceRef/ProduzioneMD con gli stati CQ/QA
            const combinedInfo = opiDetails
                ? (opiDetails + '<br><br>' + statusLines.join('<br>'))
                : statusLines.join('<br>');
            // Costruisce il box OPI completo
            opiBoxHtml = `
                    <div class="tooltip-box opi-info-tooltip">
                        <h3>Dettaglio OPI</h3>
                        <p style="color:black;">${combinedInfo}</p>
                    </div>`;
        } catch (err) {
            // Ignora eventuali errori nella lettura dei dati OPI
        }
    }

    // Prepara la sezione delle note interne da inserire all'interno del dettaglio ordine.
    // Non viene più creato un box separato, ma un blocco inline che segue le
    // altre informazioni dell'ordine.  Utilizza gli stessi colori pastello
    // definiti per le note degli arrivi.
    let noteHtml = '';
    if (taskData.noteServizio) {
        const note = String(taskData.noteServizio).trim();
        if (note) {
            noteHtml = `
                <div class="internal-note-inline">
                    <h4><span class="internal-note-icon"></span> Nota interna</h4>
                    <p>${note}</p>
                </div>`;
        }
    }
    // Costruisce l'HTML completo con i box necessari: include i dettagli ordine con le note interne
    // integrate nel primo riquadro.  Se presenti dati OPI o commenti QA, questi vengono aggiunti
    // come box separati.
    genericTooltip.innerHTML = `
        <div class="tooltip-container">
            <div class="tooltip-box shipping-info-tooltip">
                <h3>Dettaglio Ordine</h3>
                <p>${orderInfo}</p>
                ${noteHtml}
            </div>
            <div class="tooltip-box shipping-contact-tooltip"><h3>Dettaglio Cliente</h3><p>${clientInfo}</p></div>
            ${opiBoxHtml}
            ${commentBoxHtml}
        </div>`;
    
    // Mostra il tooltip completo
    showGenericTooltip('', event);
}

/**
 * Gestisce il click sull'icona del lucchetto. Chiede la password.
 */
async function handleQACommentClick(event) {
    const rowId = event.target.dataset.rowId;
    const targetRow = document.querySelector(`tr[data-row-id="${rowId}"]`);
    if (!targetRow) {
        showAlert("Errore: Riga di riferimento non trovata.");
        return;
    }

    // Crea e mostra il modale per la password
    const passwordModal = document.createElement('div');
    passwordModal.className = 'modal-overlay visible';
    passwordModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Accesso Commenti QA</h3>
            <p>Inserisci la password per modificare i commenti.</p>
            <input type="password" id="qaPasswordInput" placeholder="Password...">
            <div class="modal-buttons">
                <button id="qaConfirmBtn" class="modal-button confirm">Conferma</button>
                <button id="qaCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(passwordModal);

    const qaConfirmBtn = document.getElementById('qaConfirmBtn');
    const qaCancelBtn = document.getElementById('qaCancelBtn');
    const qaPasswordInput = document.getElementById('qaPasswordInput');

    qaConfirmBtn.onclick = () => {
        if (qaPasswordInput.value === 'qa123') {
            passwordModal.remove();
            openQACommentsEditor(targetRow);
        } else {
            showAlert("Password non corretta.");
            qaPasswordInput.value = '';
        }
    };

    qaCancelBtn.onclick = () => passwordModal.remove();
    qaPasswordInput.focus();
}

/**
 * VERSIONE AGGIORNATA - Implementa il flusso Salva -> Ricarica -> Chiudi.
 * Questa versione garantisce la coerenza visiva immediata dopo il salvataggio.
 */
function openQACommentsEditor(targetRow) {
    const currentComments = targetRow.dataset.commentiQA || '';

    // Crea e mostra la finestra di dialogo (modal) per l'editor
    const editorModal = document.createElement('div');
    editorModal.className = 'modal-overlay visible';
    // --- MODIFICA 1: Cambiato il testo del pulsante ---
    editorModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Modifica Commenti QA 🔓</h3>
            <p>Inserisci o modifica le note qui sotto.</p>
            <textarea id="qaCommentsTextarea">${currentComments}</textarea>
            <div class="modal-buttons">
                <button id="qaSaveAndUpdateBtn" class="modal-button save">Salva e Aggiorna</button>
                <button id="qaEditorCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(editorModal);

    // --- MODIFICA 2: Aggiornato l'ID del pulsante ---
    const qaSaveAndUpdateBtn = document.getElementById('qaSaveAndUpdateBtn');
    const qaEditorCancelBtn = document.getElementById('qaEditorCancelBtn');
    const qaCommentsTextarea = document.getElementById('qaCommentsTextarea');

    // --- MODIFICA 3: Riscritta completamente la logica del click ---
    qaSaveAndUpdateBtn.onclick = async () => {
        // Fornisce un feedback visivo all'utente durante le operazioni
        qaSaveAndUpdateBtn.textContent = 'Salvataggio in corso...';
        qaSaveAndUpdateBtn.disabled = true;
        qaEditorCancelBtn.disabled = true;

        // 1. Salva il nuovo commento sull'attributo della riga.
        targetRow.dataset.commentiQA = qaCommentsTextarea.value;
        
        // 2. Salva TUTTI i dati sul server.
        await saveDataToServer();
        
        // 3. Ricarica TUTTI i dati dal server per garantire la coerenza.
        qaSaveAndUpdateBtn.textContent = 'Aggiornamento dati...';
        await loadDataFromServer();
        
        // 4. Chiude la finestra di modifica solo dopo che tutto è stato completato.
        editorModal.remove();
        
        // 5. Mostra un messaggio di successo finale.
        showAlert("Commento salvato e dati ricaricati con successo!");
    };

    // La logica del pulsante "Annulla" rimane invariata
    qaEditorCancelBtn.onclick = () => editorModal.remove();
    
    // Attiva il cursore nella casella di testo
    qaCommentsTextarea.focus();
}


async function handleQACommentClick(event) {
    const rowId = event.target.dataset.rowId;
    const targetRow = document.querySelector(`tr[data-row-id="${rowId}"]`);
    if (!targetRow) {
        showAlert("Errore: Riga di riferimento non trovata.");
        return;
    }

    // Crea e mostra il modale per la password
    const passwordModal = document.createElement('div');
    passwordModal.className = 'modal-overlay visible';
    passwordModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Accesso Commenti QA</h3>
            <p>Inserisci la password per modificare i commenti.</p>
            <input type="password" id="qaPasswordInput" placeholder="Password...">
            <div class="modal-buttons">
                <button id="qaConfirmBtn" class="modal-button confirm">Conferma</button>
                <button id="qaCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(passwordModal);

    const qaConfirmBtn = document.getElementById('qaConfirmBtn');
    const qaCancelBtn = document.getElementById('qaCancelBtn');
    const qaPasswordInput = document.getElementById('qaPasswordInput');

    qaConfirmBtn.onclick = () => {
        if (qaPasswordInput.value === 'qa123') {
            passwordModal.remove();
            openQACommentsEditor(targetRow);
        } else {
            showAlert("Password non corretta.");
            qaPasswordInput.value = '';
        }
    };

    qaCancelBtn.onclick = () => passwordModal.remove();
    qaPasswordInput.focus();
}

/**
 * Apre l'editor per scrivere/modificare i commenti dopo l'inserimento della password.
 * VERSIONE MODIFICATA PER GARANTIRE LA PERSISTENZA VISIVA.
 */
/**
 * VERSIONE DEFINITIVA - Salva e forza l'aggiornamento dal server.
 * Questa è la soluzione più robusta per garantire la coerenza visiva.
 */
function openQACommentsEditor(targetRow) {
    const currentComments = targetRow.dataset.commentiQA || '';

    // Crea e mostra la finestra di dialogo (modal) per l'editor
    const editorModal = document.createElement('div');
    editorModal.className = 'modal-overlay visible';
    editorModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Modifica Commenti QA 🔓</h3>
            <p>Inserisci o modifica le note qui sotto.</p>
            <textarea id="qaCommentsTextarea">${currentComments}</textarea>
            <div class="modal-buttons">
                <button id="qaSaveAndUpdateBtn" class="modal-button save">Ricarica server e Chiudi</button>
                <button id="qaEditorCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(editorModal);

    const qaSaveAndUpdateBtn = document.getElementById('qaSaveAndUpdateBtn');
    const qaEditorCancelBtn = document.getElementById('qaEditorCancelBtn');
    const qaCommentsTextarea = document.getElementById('qaCommentsTextarea');

    // Definisce cosa succede quando si clicca su "Salva e Chiudi"
    qaSaveBtn.onclick = async () => {
        // 1. Salva il nuovo commento sull'elemento della tabella per la coerenza immediata.
        targetRow.dataset.commentiQA = qaCommentsTextarea.value;
        
        // 2. Chiude la finestra di modifica.
        editorModal.remove();
        
        // 3. Salva TUTTO sul server (questo ora funziona correttamente).
        await saveDataToServer();
        
        // 4. Mostra un messaggio che avvisa l'utente dell'aggiornamento in corso.
        showAlert("Salvataggio completato. Aggiornamento della vista in corso...");

        // 5. [LA SOLUZIONE] Attende 1 secondo per dare al server il tempo di elaborare il salvataggio.
        setTimeout(() => {
            // 6. Forza il ricaricamento di TUTTI i dati dal server, simulando il click manuale.
            //    Questo garantirà che la visualizzazione sia 100% allineata con i dati salvati.
            loadDataFromServer();
        }, 1000); // 1000 millisecondi = 1 secondo
    };

    // Chiude la finestra se si clicca "Annulla"
    qaEditorCancelBtn.onclick = () => editorModal.remove();
    
    // Attiva il cursore nella casella di testo
    qaCommentsTextarea.focus();
}

// ===================================================================
// ==> NUOVE FUNZIONI PER GESTIRE I COMMENTI QA (DA AGGIUNGERE) <==
// ===================================================================

/**
 * Gestisce il click sull'icona del lucchetto nei commenti QA.
 */
async function handleQACommentClick(event) {
    const rowId = event.target.dataset.rowId;
    const targetRow = document.querySelector(`tr[data-row-id="${rowId}"]`);
    if (!targetRow) {
        showAlert("Errore: Riga di riferimento non trovata.");
        return;
    }

    // Crea e mostra il modale per la password
    const passwordModal = document.createElement('div');
    passwordModal.className = 'modal-overlay visible';
    passwordModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Accesso Commenti QA</h3>
            <p>Inserisci la password per modificare i commenti.</p>
            <input type="password" id="qaPasswordInput" placeholder="Password...">
            <div class="modal-buttons">
                <button id="qaConfirmBtn" class="modal-button confirm">Conferma</button>
                <button id="qaCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(passwordModal);

    const qaConfirmBtn = document.getElementById('qaConfirmBtn');
    const qaCancelBtn = document.getElementById('qaCancelBtn');
    const qaPasswordInput = document.getElementById('qaPasswordInput');

    qaConfirmBtn.onclick = () => {
        if (qaPasswordInput.value === 'qa123') {
            passwordModal.remove();
            openQACommentsEditor(targetRow);
        } else {
            showAlert("Password non corretta.");
            qaPasswordInput.value = '';
        }
    };

    qaCancelBtn.onclick = () => passwordModal.remove();
    qaPasswordInput.focus();
}

/**
 * Apre l'editor per modificare i commenti QA dopo aver inserito la password corretta.
 */
function openQACommentsEditor(targetRow) {
    const currentComments = targetRow.dataset.commentiQA || '';

    // Crea e mostra il modale per l'editor dei commenti
    const editorModal = document.createElement('div');
    editorModal.className = 'modal-overlay visible';
    editorModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Modifica Commenti QA 🔓</h3>
            <p>Inserisci o modifica le note qui sotto.</p>
            <textarea id="qaCommentsTextarea">${currentComments}</textarea>
            <div class="modal-buttons">
                <button id="qaSaveBtn" class="modal-button save">Salva e Chiudi</button>
                <button id="qaEditorCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(editorModal);

    const qaSaveBtn = document.getElementById('qaSaveBtn');
    const qaEditorCancelBtn = document.getElementById('qaEditorCancelBtn');
    const qaCommentsTextarea = document.getElementById('qaCommentsTextarea');

    qaSaveAndUpdateBtn.onclick = async () => {
        // Salva il nuovo commento sull'attributo della riga
        targetRow.dataset.commentiQA = qaCommentsTextarea.value;
        // Fornisce un feedback visivo durante le operazioni
        qaSaveAndUpdateBtn.textContent = 'Aggiornamento...';
        qaSaveAndUpdateBtn.disabled = true;
        // 1. Salva tutti i dati sul server
        await saveDataToServer();
        // 2. Ricarica tutti i dati dal server per garantire la coerenza
        await loadDataFromServer();
        // 3. Chiude la finestra di modifica
        editorModal.remove();
        // 4. Aggiorna il Gantt di magazzino per visualizzare subito il commento
        updateWarehouseGanttChart();
        // 5. Mostra un messaggio di successo
        showAlert("Commento salvato e server ricaricato con successo!");
    };

    qaEditorCancelBtn.onclick = () => editorModal.remove();
    qaCommentsTextarea.focus();
}


function showGenericTooltip(htmlContent, event) {
        if (!genericTooltip) genericTooltip = document.getElementById('genericTooltip');
        
        // Se htmlContent ha del testo, usa la vecchia logica (tooltip semplice nero).
        // Se è vuoto, mostra il contenuto già preparato dalle funzioni "split".
        if (htmlContent) {
            genericTooltip.innerHTML = htmlContent;
            // Ripristina lo stile base per il tooltip semplice
            genericTooltip.style.backgroundColor = 'rgba(0, 0, 0, 0.85)';
            genericTooltip.style.padding = '10px 15px';
            genericTooltip.style.border = 'none';
            genericTooltip.style.color = 'white';
            genericTooltip.style.maxWidth = '350px';
        }

        genericTooltip.classList.add('visible');
        // Posiziona subito il tooltip in base all'evento iniziale
        moveGenericTooltip(event);
        // Aggiunge un listener per seguire il movimento del mouse mentre il
        // tooltip è visibile.  Questo permette all'utente di vedere
        // l'elemento sottostante senza che il tooltip copra il puntatore.
        document.addEventListener('mousemove', moveGenericTooltip);
    }

function hideGenericTooltip() {
    if (genericTooltip) {
        genericTooltip.classList.remove('visible');
    }
    // Quando il tooltip viene nascosto, rimuove il listener del mouse per
    // evitare overhead inutile.
    document.removeEventListener('mousemove', moveGenericTooltip);
}

// Gestisce il click sul pallino CQ: chiede la password e apre l'editor dello stato CQ.
// L'operatore CQ può cambiare il colore del pallino (white, yellow, green, red)
// solo dopo aver inserito la password corretta (cq456). Il nuovo stato viene
// salvato e ricaricato dal server per garantire coerenza immediata.
async function handleCQStatusClick(event) {
    const rowId = event.target.dataset.rowId;
    const targetRow = document.querySelector(`tr[data-row-id="${rowId}"]`);
    if (!targetRow) {
        showAlert("Errore: Riga di riferimento non trovata.");
        return;
    }
    // Crea e mostra il modale per la password CQ
    const passwordModal = document.createElement('div');
    passwordModal.className = 'modal-overlay visible';
    passwordModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Accesso CQ</h3>
            <p>Inserisci la password per modificare lo stato CQ.</p>
            <input type="password" id="cqPasswordInput" placeholder="Password...">
            <div class="modal-buttons">
                <button id="cqConfirmBtn" class="modal-button confirm">Conferma</button>
                <button id="cqCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(passwordModal);
    const cqConfirmBtn = document.getElementById('cqConfirmBtn');
    const cqCancelBtn = document.getElementById('cqCancelBtn');
    const cqPasswordInput = document.getElementById('cqPasswordInput');
    cqConfirmBtn.onclick = () => {
        if (cqPasswordInput.value === 'cq456') {
            passwordModal.remove();
            openCQStatusEditor(targetRow);
        } else {
            showAlert("Password non corretta.");
            cqPasswordInput.value = '';
        }
    };
    cqCancelBtn.onclick = () => passwordModal.remove();
    cqPasswordInput.focus();
}

// Gestisce il click sulla bandierina QA: chiede la password e apre l'editor dello stato QA.
// L'operatore QA può cambiare il colore della bandierina (white, yellow, green, red)
// solo dopo aver inserito la password corretta (qa123). Il nuovo stato viene
// salvato e ricaricato dal server per garantire coerenza immediata.
async function handleQAStatusClick(event) {
    const rowId = event.target.dataset.rowId;
    const targetRow = document.querySelector(`tr[data-row-id="${rowId}"]`);
    if (!targetRow) {
        showAlert("Errore: Riga di riferimento non trovata.");
        return;
    }
    // Crea e mostra il modale per la password QA
    const passwordModal = document.createElement('div');
    passwordModal.className = 'modal-overlay visible';
    passwordModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Accesso QA</h3>
            <p>Inserisci la password per modificare lo stato QA.</p>
            <input type="password" id="qaPasswordInput" placeholder="Password...">
            <div class="modal-buttons">
                <button id="qaConfirmBtn" class="modal-button confirm">Conferma</button>
                <button id="qaCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(passwordModal);
    const qaConfirmBtn = document.getElementById('qaConfirmBtn');
    const qaCancelBtn = document.getElementById('qaCancelBtn');
    const qaPasswordInput = document.getElementById('qaPasswordInput');
    qaConfirmBtn.onclick = () => {
        if (qaPasswordInput.value === 'qa123') {
            passwordModal.remove();
            openQAStatusEditor(targetRow);
        } else {
            showAlert("Password non corretta.");
            qaPasswordInput.value = '';
        }
    };
    qaCancelBtn.onclick = () => passwordModal.remove();
    qaPasswordInput.focus();
}

// Mostra l'editor per scegliere il nuovo stato QA. Solo dopo la password corretta.
function openQAStatusEditor(targetRow) {
    const currentStatus = targetRow.dataset.qaStatus || 'white';
    const editorModal = document.createElement('div');
    editorModal.className = 'modal-overlay visible';
    editorModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Modifica Stato QA</h3>
            <p>Seleziona il nuovo stato per la merce.</p>
            <div class="qa-status-options">
                <label><input type="radio" name="qaStatusOption" value="white" ${currentStatus === 'white' ? 'checked' : ''}><span class="qa-status-flag qa-status-white"></span> Merce in fase di valutazione</label><br>
                <label><input type="radio" name="qaStatusOption" value="yellow" ${currentStatus === 'yellow' ? 'checked' : ''}><span class="qa-status-flag qa-status-yellow"></span> Merce accettata con deroga</label><br>
                <label><input type="radio" name="qaStatusOption" value="green" ${currentStatus === 'green' ? 'checked' : ''}><span class="qa-status-flag qa-status-green"></span> Merce conforme</label><br>
                <label><input type="radio" name="qaStatusOption" value="red" ${currentStatus === 'red' ? 'checked' : ''}><span class="qa-status-flag qa-status-red"></span> Merce non conforme</label>
            </div>
            <div class="modal-buttons">
                <button id="qaStatusSaveBtn" class="modal-button save">Salva</button>
                <button id="qaStatusCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(editorModal);
    const saveBtn = document.getElementById('qaStatusSaveBtn');
    const cancelBtn = document.getElementById('qaStatusCancelBtn');
    saveBtn.onclick = async () => {
        const selected = editorModal.querySelector('input[name="qaStatusOption"]:checked');
        if (selected) {
            // Aggiorna lo stato sul dataset della riga
            const newVal = selected.value;
            targetRow.dataset.qaStatus = newVal;
            // Aggiorna immediatamente il colore della bandierina QA nella riga
            const qaFlag = targetRow.querySelector('.qa-status-flag');
            if (qaFlag) {
                // Rimuovi tutte le possibili classi di stato
                qaFlag.classList.remove('qa-status-white', 'qa-status-yellow', 'qa-status-green', 'qa-status-red');
                qaFlag.classList.add(`qa-status-${newVal}`);
            }
            // Registra il cambiamento per notificare gli altri utenti
            if (typeof registerQualityStatusChange === 'function') {
                registerQualityStatusChange('QA', targetRow);
            }
            editorModal.remove();
            // Salva tutte le modifiche sul server e ricarica i dati per rendere
            // immediatamente visibile il nuovo stato nel Gantt e nelle tabelle
            await saveDataToServer();
            await loadDataFromServer();
            updateWarehouseGanttChart();
            showAlert("Stato QA aggiornato con successo!");
            // Dopo il salvataggio, verifica se ci sono avvisi da mostrare
            if (typeof checkAndNotifyQuality === 'function') {
                checkAndNotifyQuality();
            }

    // Rende trascinabili i pop‑up di notifica (ADR e CQ/QA).  Questo
    // evita che i bubble rimangano al centro dello schermo e impediscano
    // la vista delle tabelle.  L'utente può trascinare l'avviso con il
    // mouse e posizionarlo dove preferisce.  La funzione si attiva
    // automaticamente per gli avvisi esistenti.
    function makeNotificationDraggable(el) {
        if (!el) return;
        let offsetX = 0;
        let offsetY = 0;
        let isDown = false;
        el.addEventListener('mousedown', (ev) => {
            // evita di trascinare quando si cliccano pulsanti o link
            const tag = ev.target.tagName.toLowerCase();
            if (tag === 'button' || tag === 'a' || ev.target.classList.contains('adr-close-btn') || ev.target.classList.contains('quality-close-btn')) {
                return;
            }
            isDown = true;
            const rect = el.getBoundingClientRect();
            offsetX = ev.clientX - rect.left;
            offsetY = ev.clientY - rect.top;
            function moveHandler(e) {
                if (!isDown) return;
                el.style.left = `${e.clientX - offsetX}px`;
                el.style.top = `${e.clientY - offsetY}px`;
            }
            function upHandler() {
                isDown = false;
                document.removeEventListener('mousemove', moveHandler);
                document.removeEventListener('mouseup', upHandler);
            }
            document.addEventListener('mousemove', moveHandler);
            document.addEventListener('mouseup', upHandler);
        });
    }
    // Applica la funzionalità di drag ai pop‑up se sono presenti nel DOM
    makeNotificationDraggable(document.getElementById('qualityNotification'));
    makeNotificationDraggable(document.getElementById('adrNotification'));
    // Rende trascinabile anche l'avviso di spedizione
    makeNotificationDraggable(document.getElementById('shippingNotification'));
        }
    };
    cancelBtn.onclick = () => editorModal.remove();
}

// Mostra l'editor per scegliere il nuovo stato CQ. Solo dopo la password corretta.
function openCQStatusEditor(targetRow) {
    const currentStatus = targetRow.dataset.cqStatus || 'white';
    const editorModal = document.createElement('div');
    editorModal.className = 'modal-overlay visible';
    editorModal.innerHTML = `
        <div class="modal-content qa-modal-content">
            <h3>Modifica Stato CQ</h3>
            <p>Seleziona il nuovo stato per la merce.</p>
            <div class="cq-status-options">
                <label><input type="radio" name="cqStatusOption" value="white" ${currentStatus === 'white' ? 'checked' : ''}><span class="cq-status-dot cq-status-white"></span> Merce da analizzare</label><br>
                <label><input type="radio" name="cqStatusOption" value="yellow" ${currentStatus === 'yellow' ? 'checked' : ''}><span class="cq-status-dot cq-status-yellow"></span> Merce accettata con deroga</label><br>
                <label><input type="radio" name="cqStatusOption" value="green" ${currentStatus === 'green' ? 'checked' : ''}><span class="cq-status-dot cq-status-green"></span> Merce conforme</label><br>
                <label><input type="radio" name="cqStatusOption" value="red" ${currentStatus === 'red' ? 'checked' : ''}><span class="cq-status-dot cq-status-red"></span> Merce non conforme</label>
            </div>
            <div class="modal-buttons">
                <button id="cqStatusSaveBtn" class="modal-button save">Salva</button>
                <button id="cqStatusCancelBtn" class="modal-button cancel">Annulla</button>
            </div>
        </div>
    `;
    document.body.appendChild(editorModal);
    const saveBtn = document.getElementById('cqStatusSaveBtn');
    const cancelBtn = document.getElementById('cqStatusCancelBtn');
    saveBtn.onclick = async () => {
        const selected = editorModal.querySelector('input[name="cqStatusOption"]:checked');
        if (selected) {
            // Aggiorna lo stato sul dataset della riga
            const newVal = selected.value;
            targetRow.dataset.cqStatus = newVal;
            // Aggiorna immediatamente il colore del pallino CQ nella riga
            const cqDot = targetRow.querySelector('.cq-status-dot');
            if (cqDot) {
                cqDot.classList.remove('cq-status-white', 'cq-status-yellow', 'cq-status-green', 'cq-status-red');
                cqDot.classList.add(`cq-status-${newVal}`);
            }
            // Registra il cambiamento per notificare gli altri utenti
            if (typeof registerQualityStatusChange === 'function') {
                registerQualityStatusChange('CQ', targetRow);
            }
            editorModal.remove();
            // Salva tutte le modifiche sul server e ricarica i dati per
            // rendere immediatamente visibile il nuovo stato
            await saveDataToServer();
            await loadDataFromServer();
            updateWarehouseGanttChart();
            showAlert("Stato CQ aggiornato con successo!");
            // Dopo il salvataggio, verifica se ci sono avvisi da mostrare
            if (typeof checkAndNotifyQuality === 'function') {
                checkAndNotifyQuality();
            }
        }
    };
    cancelBtn.onclick = () => editorModal.remove();
}
function moveGenericTooltip(event) {
    if (!genericTooltip || !genericTooltip.classList.contains('visible') || !event) return;
    const tooltipRect = genericTooltip.getBoundingClientRect();
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;
    // Calcola la posizione orizzontale centrando il tooltip attorno al puntatore.
    // Questa logica impedisce al tooltip di uscire a destra o a sinistra del viewport,
    // spostandolo in modo equilibrato rispetto al punto di origine.
    let x = event.clientX - (tooltipRect.width / 2);
    // Impedisci che il tooltip vada oltre il margine sinistro.
    if (x < 10) x = 10;
    // Impedisci che il tooltip vada oltre il margine destro.
    if (x + tooltipRect.width > viewportWidth - 10) {
        x = viewportWidth - tooltipRect.width - 10;
    }
    // Posizione verticale: per default sotto il puntatore con offset.
    let y = event.clientY + 20;
    // Se il tooltip esce dal fondo dello schermo, posizionalo sopra il puntatore.
    if (y + tooltipRect.height > viewportHeight - 10) {
        y = event.clientY - tooltipRect.height - 20;
    }
    // Impedisci di finire sopra il bordo superiore.
    if (y < 10) y = 10;
    genericTooltip.style.left = `${x}px`;
    genericTooltip.style.top = `${y}px`;
}


    const ganttDayHeaderColors = [
        '#E0F2F7', '#F8E7F0', '#E6F0E6', '#F7F7E0', '#E0E7F7',
        '#F0E0F7', '#E7F7F0', '#F7F0E0', '#E0F7F0', '#F0F7E0',
        '#E7E0F7', '#F7E0F0', '#E0E0F7', '#F0F0E0'
    ];

    function updateGanttChart() {
        ganttChartDiv.innerHTML = '';

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        // Imposta un intervallo di 30 giorni per il Gantt di produzione invece dei 14 originari.
        const dateCount = 30;
        const dates = Array.from({ length: dateCount }, (_, i) => {
            const d = new Date(today);
            d.setDate(today.getDate() + i);
            return d;
        });

        // Imposta la griglia in base al numero di colonne; la prima colonna rimane fissa per le etichette dei macchinari.
        if (ganttChartDiv) {
            ganttChartDiv.style.display = 'grid';
            ganttChartDiv.style.gridTemplateColumns = '200px repeat(' + dateCount + ', 110px)';
            ganttChartDiv.style.minWidth = (200 + dateCount * 110) + 'px';
        }
        const emptyHeader = document.createElement('div');
        emptyHeader.classList.add('gantt-header', 'gantt-row-header');
        emptyHeader.textContent = 'Macchinario / Ordini';
        ganttChartDiv.appendChild(emptyHeader);

        dates.forEach((date, index) => {
            const dateHeader = document.createElement('div');
            dateHeader.classList.add('gantt-header');
            dateHeader.style.backgroundColor = ganttDayHeaderColors[index % ganttDayHeaderColors.length];
            const dayOfWeek = date.toLocaleDateString('it-IT', { weekday: 'long' });
            dateHeader.innerHTML = `${date.toLocaleDateString('it-IT', { day: '2-digit', month: 'short' })}<br><span class="day-of-week">${dayOfWeek}</span>`;
            if (date.getDay() === 0 || date.getDay() === 6) dateHeader.classList.add('weekend');
            ganttChartDiv.appendChild(dateHeader);
        });

        const allRowsData = Array.from(productionTableBody.querySelectorAll('tr')).map(row => getRowData(row));
        /*
         * Integra le informazioni mancanti (OP, OV, lotto, operatore) nel dataset
         * utilizzato per costruire il Gantt della produzione.  Per ogni riga
         * della tabella di produzione cerchiamo una riga corrispondente nella
         * tabella "Programma Giornaliero di Produzione" (dailyProductionTable)
         * basandoci sul codice articolo.  Se trovata, copiamo i campi op, ov,
         * lotto e operatore dal programma giornaliero nella riga di
         * produzione, evitando di sovrascrivere valori già presenti.  In
         * questo modo il Gantt e il relativo tooltip possono mostrare tali
         * informazioni aggiuntive.
         */
        try {
            const dailyRows = Array.from(dailyProductionTableBody.querySelectorAll('tr')).map(dr => getDailyRowData(dr));
            allRowsData.forEach(row => {
                const match = dailyRows.find(d => {
                    const c1 = String(d.codice || '').trim().toUpperCase();
                    const c2 = String(row.codice || '').trim().toUpperCase();
                    return c1 && c1 === c2;
                });
                if (match) {
                    if (!row.op || row.op.trim() === '') row.op = match.op || '';
                    if (!row.ov || row.ov.trim() === '') row.ov = match.ov || '';
                    if (!row.lotto || row.lotto.trim() === '') row.lotto = match.lotto || '';
                    // Se l'operatore della tabella di produzione è vuoto, usa quello del programma giornaliero
                    if (!row.operatore || row.operatore.trim() === '') {
                        row.operatore = match.operatore || row.operatore;
                    }
                }
            });
        } catch (e) {
            console.warn('Impossibile associare i dati del programma giornaliero al Gantt di produzione:', e);
        }

        /*
         * Crea una mappa che associa (codice + data di produzione) alle informazioni
         * complete (OP, OV, lotto, operatore) della riga di produzione.  Questa mappa
         * viene poi utilizzata per arricchire le righe di confezionamento che hanno
         * la stessa data di confezionamento e codice ma non possiedono tutti i
         * campi compilati.  In questo modo il Gantt della confezione mostra più
         * informazioni quando possibile.
         */
        const productionInfoByKey = {};
        allRowsData.forEach(row => {
            const codeKey = String(row.codice || '').trim().toUpperCase();
            const prodDate = String(row.produzioneData || '').trim();
            if (!codeKey || !prodDate) return;
            const key = `${codeKey}_${prodDate}`;
            if (!productionInfoByKey[key]) {
                productionInfoByKey[key] = {
                    op: row.op || '',
                    ov: row.ov || '',
                    lotto: row.lotto || '',
                    operatore: row.operatore || ''
                };
            }
        });
        // Arricchisci le righe di confezionamento con i dati mancanti
        allRowsData.forEach(row => {
            const codeKey = String(row.codice || '').trim().toUpperCase();
            const packDate = String(row.dataConfezionamento || '').trim();
            if (!codeKey || !packDate) return;
            const key = `${codeKey}_${packDate}`;
            const info = productionInfoByKey[key];
            if (info) {
                if ((!row.op || row.op.trim() === '') && info.op) row.op = info.op;
                if ((!row.ov || row.ov.trim() === '') && info.ov) row.ov = info.ov;
                if ((!row.lotto || row.lotto.trim() === '') && info.lotto) row.lotto = info.lotto;
                if ((!row.operatore || row.operatore.trim() === '') && info.operatore) row.operatore = info.operatore;
            }
        });

        const uniqueMachines = new Set(macchinariOptions);
        allRowsData.forEach(row => {
            if (row.macchinari && row.macchinari.trim() !== '') uniqueMachines.add(row.macchinari.trim());
        });
        const machinesUsedInGantt = Array.from(uniqueMachines).sort();

        machinesUsedInGantt.forEach(machine => {
            const machineHeader = document.createElement('div');
            machineHeader.classList.add('gantt-cell', 'gantt-row-header');
            machineHeader.textContent = machine;
            machineHeader.style.backgroundColor = '#e9ecef';
            ganttChartDiv.appendChild(machineHeader);

            dates.forEach(date => {
                const cell = document.createElement('div');
                cell.classList.add('gantt-cell');
                if (date.getDay() === 0 || date.getDay() === 6) cell.classList.add('weekend');

                const tasksForCell = allRowsData.filter(row => {
                    const prodDateParts = row.produzioneData.split('/');
                    const prodDate = prodDateParts.length === 3 ? new Date(parseInt(prodDateParts[2]), parseInt(prodDateParts[1]) - 1, parseInt(prodDateParts[0])) : null;
                    if (!prodDate || isNaN(prodDate.getTime()) || !row.macchinari || row.macchinari !== machine || !(parseFloat(row.quantitaDaProdurre) > 0)) return false;

                    prodDate.setHours(0, 0, 0, 0);
                    const daysOfProduction = parseInt(row.giorniDiProduzione) || 1;
                    const endDate = new Date(prodDate);
                    endDate.setDate(prodDate.getDate() + daysOfProduction - 1);
                    return date >= prodDate && date <= endDate;
                });

                tasksForCell.forEach(task => {
                    const taskElement = document.createElement('div');
                    let isMedical = isMedicalDeviceCode(task.codice);

                    if (isMedical) {
                        taskElement.classList.add('gantt-task', 'production-4xxxx');
                    } else {
                        taskElement.classList.add('gantt-task', 'production-task');
                    }
                    if (task.materiePrime === 'si') taskElement.classList.add('materie-si');
                    else if (task.materiePrime === 'no') taskElement.classList.add('materie-no');

                    // Visualizza anche OP e OV in maniera compatta accanto al codice per
                    // fornire un'informazione immediata all'operatore.  Se non sono
                    // presenti vengono omessi.  Il resto dei dettagli (lotto e operatore)
                    // verrà mostrato nel tooltip.
                    let codeInfo = '';
                    if (task.op && task.ov) {
                        codeInfo = `OP: ${task.op} / OV: ${task.ov}`;
                    } else if (task.op) {
                        codeInfo = `OP: ${task.op}`;
                    } else if (task.ov) {
                        codeInfo = `OV: ${task.ov}`;
                    }
                    const codiceLabel = task.codice ? `Codice: ${task.codice}` : '';
                    taskElement.innerHTML = `<span class="task-code">${codeInfo ? codeInfo + ' - ' : ''}${codiceLabel}</span><span class="task-details">Prodotto: ${task.prodotto}</span>`;
                    taskElement.addEventListener('mouseover', (e) => showSplitTooltip(task, e));
                    taskElement.addEventListener('mouseout', hideGenericTooltip);
                    cell.appendChild(taskElement);
                });

                ganttChartDiv.appendChild(cell);
            });
        });

        const packagingHeader = document.createElement('div');
        packagingHeader.classList.add('gantt-cell', 'gantt-row-header');
        packagingHeader.textContent = 'Confezionamento';
        packagingHeader.style.backgroundColor = '#e9ecef';
        ganttChartDiv.appendChild(packagingHeader);
        dates.forEach(date => {
            const cell = document.createElement('div');
            cell.classList.add('gantt-cell');
            const dateKey = date.toISOString().split('T')[0];
            if (date.getDay() === 0 || date.getDay() === 6) cell.classList.add('weekend');

            const packagingTasks = allRowsData.filter(row => {
                const packDateParts = row.dataConfezionamento.split('/');
                const packDate = packDateParts.length === 3 ? new Date(parseInt(packDateParts[2]), parseInt(packDateParts[1]) - 1, parseInt(packDateParts[0])) : null;
                return packDate && !isNaN(packDate.getTime()) && packDate.toISOString().split('T')[0] === dateKey;
            });
                packagingTasks.forEach(task => {
                const taskElement = document.createElement('div');
                let isMedical = isMedicalDeviceCode(task.codice);

                if (isMedical) {
                    taskElement.classList.add('gantt-task', 'packaging-4xxxx');
                } else {
                    taskElement.classList.add('gantt-task', 'packaging-task');
                }
                if (task.materialeConfezionamento === 'si') taskElement.classList.add('materie-si');
                else if (task.materialeConfezionamento === 'no') taskElement.classList.add('materie-no');

                // Visualizza anche OP e OV accanto al codice per una visione rapida delle
                // principali informazioni di produzione legate al confezionamento.  Se
                // queste informazioni non sono disponibili vengono omesse.  Lotto e
                // operatore saranno visualizzati nel tooltip.
                let codeInfo = '';
                if (task.op && task.ov) {
                    codeInfo = `OP: ${task.op} / OV: ${task.ov}`;
                } else if (task.op) {
                    codeInfo = `OP: ${task.op}`;
                } else if (task.ov) {
                    codeInfo = `OV: ${task.ov}`;
                }
                const codiceLabel = task.codice ? `Codice: ${task.codice}` : '';
                taskElement.innerHTML = `<span class="task-code">${codeInfo ? codeInfo + ' - ' : ''}${codiceLabel}</span><span class="task-details">Prodotto: ${task.prodotto}</span>`;
                taskElement.addEventListener('mouseover', (e) => showSplitTooltip(task, e));
                taskElement.addEventListener('mouseout', hideGenericTooltip);
                cell.appendChild(taskElement);
            });
            ganttChartDiv.appendChild(cell);
        });
    }

   
function updateWarehouseGanttChart() {
    warehouseGanttChartDiv.innerHTML = '';
    const ganttGrid = document.createElement('div');
    ganttGrid.className = 'gantt-chart warehouse-gantt-chart';
    // Imposta la larghezza fissa delle colonne del Gantt di magazzino. Ogni colonna
    // rappresenta un giorno e viene impostata a circa 110px.  Ora consideriamo
    // un intervallo di 30 giorni e lasciamo all'utente la possibilità di
    // scorrere orizzontalmente per vedere le colonne oltre le prime 14 visibili.
    ganttGrid.style.gridTemplateColumns = '200px repeat(30, 110px)';
    const dates = Array.from({ length: 30 }, (_, i) => {
        const d = new Date();
        d.setHours(12, 0, 0, 0);
        d.setDate(d.getDate() + i);
        return d;
    });

    ganttGrid.appendChild(document.createElement('div'));
    dates.forEach((date, index) => {
        const dateHeader = document.createElement('div');
        dateHeader.classList.add('gantt-header');
        dateHeader.style.backgroundColor = ganttDayHeaderColors[index % ganttDayHeaderColors.length];
        const dayOfWeek = date.toLocaleDateString('it-IT', { weekday: 'long' });
        dateHeader.innerHTML = `${date.toLocaleDateString('it-IT', { day: '2-digit', month: 'short' })}<br><span class="day-of-week">${dayOfWeek}</span>`;
        if (date.getDay() === 0 || date.getDay() === 6) dateHeader.classList.add('weekend');
        ganttGrid.appendChild(dateHeader);
    });

    // Mappatura per raggruppare più famiglie dello stesso tipo sotto un unico colore.
    // Le famiglie elencate qui (in minuscolo) condivideranno la stessa classe colore.
    const unifiedFamiliesMap = {
        'cespiti': 'gruppo-servizi',
        'cespite': 'gruppo-servizi',
        'servizi': 'gruppo-servizi',
        'servizi vari': 'gruppo-servizi',
        'servizi vari/miscelati': 'gruppo-servizi',
        'attrezzatura varia': 'gruppo-servizi',
        'attrezzatura': 'gruppo-servizi',
        'campioni di laboratorio': 'gruppo-servizi',
        'campioni laboratorio': 'gruppo-servizi',
        'materiale di consumo laboratorio': 'gruppo-servizi',
        'materiale consumo laboratorio': 'gruppo-servizi',
        'materiale laboratorio': 'gruppo-servizi'
    };
    const familyColorMap = new Map();
    let colorIndex = 0;
    const getFamilyColorClass = (familyName) => {
        if (!familyName) familyName = 'Senza Famiglia';
        const lower = String(familyName || '').toLowerCase().trim();
        const groupKey = unifiedFamiliesMap[lower] || lower;
        if (!familyColorMap.has(groupKey)) {
            familyColorMap.set(groupKey, `family-color-${colorIndex % 8}`);
            colorIndex++;
        }
        return familyColorMap.get(groupKey);
    };

    const specialBorderCodes = new Set(['BEC0305', 'BEC0506', 'BEC0706', 'BEC0810', 'BEC0910', 'BEC1010', 'BEC1111','BEC1214', 'BEC1415', 'BEC1818', 'BEC1918', 'BEC2420', 'BEC2520', 'BEC2620','BEC3024', 'STV0214', 'STV0314', 'STV1014', 'STV1114', 'STV2420', 'BOR1321','CAP0106', 'CAP0308', 'CAP0408', 'CAP0513', 'CAP0613']);

    /*
     * Elenco dei codici articolo che richiedono il trasporto ADR.
     * Questa lista è ricavata dal file “ADR_technics.xlsx” fornito dall’utente.  Durante la
     * generazione del Gantt viene controllato se un codice appartiene a questo set: se sì,
     * viene visualizzato un indicatore lampeggiante “ADR” sulla riga di spedizione e
     * nel tooltip viene aggiunta un’avvertenza in rosso.  Tutti i codici sono
     * convertiti in maiuscolo per semplificare la corrispondenza.
     */
    const adrCodes = new Set([
        '10004', '10004-0,05', '10004-0,5', '10004-1', '10004-10', '10004-5',
        '1385', '1385-0,05', '9989', '9989-5',
        '9660', '9660*', '9660-0,05', '9660-0,5', '9660-1', '9660-10', '9660-5',
        '9298', '9298-0,05',
        '9299B', '9299B-0,05', '9299B-0,5', '9299B-1', '9299B-10', '9299B-25', '9299B-30', '9299B-5',
        '3015', '3015-0,05', '3015-0,5', '3015-1', '3015-25', '3015-5', '3015SM',
        '18702', '18702-1',
        '9707', '9707-0,05', '9707-25',
        '1449', '1449-0,05'
    ].map(code => code.toUpperCase()));
    // Rende la lista ADR disponibile globalmente, così da poter essere
    // utilizzata anche in altre funzioni (es. tooltip).  Non altera la
    // variabile locale, ma crea una proprietà su window per accesso globale.
    window.adrCodes = adrCodes;

    const renderSectionRow = (title, data, sectionClass, isArrivalSection = false) => {
        const rowHeader = document.createElement('div');
        rowHeader.className = 'gantt-row-header';
        if (isArrivalSection) {
            const uniqueFamilies = [...new Set(data.map(item => item.family || 'Senza Famiglia'))].sort();
            uniqueFamilies.forEach(family => getFamilyColorClass(family));
            let legendHtml = '<div class="gantt-legend">';
            uniqueFamilies.forEach(family => {
                const colorClass = getFamilyColorClass(family);
                const tempTask = document.createElement('div');
                tempTask.className = `gantt-task ${colorClass}`;
                tempTask.style.display = 'none';
                document.body.appendChild(tempTask);
                const bgColor = window.getComputedStyle(tempTask).backgroundColor;
                document.body.removeChild(tempTask);
                legendHtml += `<div class="legend-item"><div class="legend-color-box" style="background-color: ${bgColor};"></div><span class="legend-text">${family}</span></div>`;
            });
            legendHtml += '</div>';
            // Costruisce la legenda per il magazzino.  Include lo stato bianco
            // (merce da evadere) e lo stato verde (merce evasa).  La legenda viene
            // mostrata solo nella sezione arrivi.
            const magLegendHtml = '<div class="mag-legend"><span class="mag-legend-title">Legenda Magazzino:</span>' +
                '<div class="mag-legend-item"><span class="mag-status-dot mag-status-white"></span><span> Merce da evadere</span></div>' +
                '<div class="mag-legend-item"><span class="mag-status-dot mag-status-green"></span><span> Merce evasa</span></div>' +
            '</div>';
            rowHeader.innerHTML = `<strong>${title}</strong>${legendHtml}${magLegendHtml}`;
        } else {
            // Aggiunge le legende CQ e QA accanto al titolo per la sezione spedizioni.
            const cqLegendHtml = '<div class="cq-legend"><span class="cq-legend-title">Legenda CQ:</span>' +
                '<div class="cq-legend-item"><span class="cq-status-dot cq-status-white"></span><span> Merce da analizzare / in analisi</span></div>' +
                '<div class="cq-legend-item"><span class="cq-status-dot cq-status-yellow"></span><span> Merce accettata con deroga</span></div>' +
                '<div class="cq-legend-item"><span class="cq-status-dot cq-status-green"></span><span> Merce conforme</span></div>' +
                '<div class="cq-legend-item"><span class="cq-status-dot cq-status-red"></span><span> Merce non conforme</span></div>' +
            '</div>';
            const qaLegendHtml = '<div class="qa-legend"><span class="qa-legend-title">Legenda QA:</span>' +
                '<div class="qa-legend-item"><span class="qa-status-flag qa-status-white"></span><span> Merce in fase di valutazione</span></div>' +
                '<div class="qa-legend-item"><span class="qa-status-flag qa-status-yellow"></span><span> Merce accettata con deroga</span></div>' +
                '<div class="qa-legend-item"><span class="qa-status-flag qa-status-green"></span><span> Merce conforme</span></div>' +
                '<div class="qa-legend-item"><span class="qa-status-flag qa-status-red"></span><span> Merce non conforme</span></div>' +
            '</div>';
            // Per la sezione spedizioni inseriamo le legende CQ e QA una sotto l'altra,
            // aggiungendo uno spazio verticale dedicato tra di esse. In questo modo
            // l'utente percepisce chiaramente che si tratta di due blocchi distinti.
            rowHeader.innerHTML = `<strong>${title}</strong>${cqLegendHtml}<div class="legend-separator"></div>${qaLegendHtml}`;
        }
        ganttGrid.appendChild(rowHeader);

        dates.forEach(date => {
            const cell = document.createElement('div');
            cell.classList.add('gantt-cell');
            if (date.getDay() === 0 || date.getDay() === 6) cell.classList.add('weekend');
            const dateKey = date.toLocaleDateString('it-IT');
            const tasksForDay = data.filter(row => row.dataConsegna === dateKey);
            const ovGroups = tasksForDay.reduce((acc, task) => {
                const ov = task.ov || 'Senza OV';
                const family = task.family || 'Senza Famiglia';
                const groupKey = `${ov}-${family}`;
                if (!acc[groupKey]) {
                    acc[groupKey] = { ov: ov, family: family, tasks: [] };
                }
                acc[groupKey].tasks.push(task);
                return acc;
            }, {});

            for (const key in ovGroups) {
                const group = ovGroups[key];
                const groupContainer = document.createElement('div');
                groupContainer.className = 'gantt-ov-group';
                if (sectionClass) groupContainer.classList.add(sectionClass);
                const groupHeader = document.createElement('div');
                groupHeader.className = 'gantt-ov-group-header';
                const firstTaskInGroup = group.tasks[0];
                const layoutIconHtml = getLayoutIcon(firstTaskInGroup.layout);
                
                // ==========================================================
                // ==> INIZIO BLOCCO DI LOGICA UNIFICATO (da V1.56 e V1.57) <==
                // ==========================================================
                let priorityIconHtml = '';
                let starIconHtml = '';

                if (isArrivalSection) {
                    // Logica per le stelle (gialla e blu)
                    const medicalDeviceCodes = ['BEC0305', 'PH701/50C', 'BEC0506', 'BEC0706', 'BEC0810', 'BEC0910', 'BEC1010', 'BEC1111', 'BEC1214', 'BEC1415', 'BEC1818', 'BEC1918', 'BEC2420', 'BEC2520', 'BEC2620', 'BEC3024', 'STV0214', 'STV0314', 'STV1014', 'STV1114', 'STV2420', 'BOR1321', 'CAP0106', 'CAP0308', 'CAP0408', 'CAP0513', 'CAP0613'];
                    const hasYellowStar = group.tasks.some(task => medicalDeviceCodes.includes(String(task.codiceArticolo || '').toUpperCase()) || String(task.descrizioneArticolo || '').toLowerCase().includes('ago'));
                    
                    if (hasYellowStar) {
                        starIconHtml = '<span class="gantt-star-icon yellow-star">★</span>';
                    } else {
                        const hasBlueStar = group.tasks.some(task => String(task.codiceArticolo || '').toUpperCase().startsWith('PIL') || String(task.codiceArticolo || '').toUpperCase().startsWith('EGC'));
                        if (hasBlueStar) {
                            starIconHtml = '<span class="gantt-star-icon blue-star">★</span>';
                        }
                    }

                    // Logica per il pallino rosso luminoso (priorità)
                    const priorityCodes = ['71866A', '71866B', '71866C', '71866D', '71360B', '71360A'];
                    const priorityKeywords = ['sodium hyaluronate', 'sodio ialuronato', 'hpdr na', 'hprna', 'acido ialuronico'];
                    const hasPriority = group.tasks.some(task => {
                        const layout = String(task.layout || '').toUpperCase();
                        const code = String(task.codiceArticolo || '').toUpperCase();
                        const description = String(task.descrizioneArticolo || '').toLowerCase();
                        return layout.includes('G5CELLA +4°C') && (priorityCodes.includes(code) || priorityKeywords.some(keyword => description.includes(keyword)));
                    });
                    if (hasPriority) {
                        priorityIconHtml = '<span class="priority-icon"></span>';
                    }
                }
                
                // Prima di mostrare il gruppo delle spedizioni, filtra le
                // righe che rappresentano servizi (ad esempio "DESCRIZIONI SERVIZI E VARIE" o
                // "MARCA DA BOLLO").  Per le spedizioni, se dopo il filtro non restano
                // task da visualizzare, salta completamente il gruppo per evitare
                // la visualizzazione di un OV vuoto.  Per gli arrivi non è
                // applicato alcun filtro (tutte le righe vengono mostrate).
                // Filtra i task da visualizzare.  Per la sezione degli arrivi
                // escludi le righe già evase (magStatus === 'green'); per le
                // spedizioni filtra le righe di servizio (es. "DESCRIZIONI SERVIZI E VARIE" e "MARCA DA BOLLO").
                const tasksToUse = isArrivalSection
                    ? group.tasks.filter(t => t.magStatus !== 'green')
                    : group.tasks.filter(t => {
                        const d = String(t.descrizioneArticolo || '').trim().toUpperCase();
                        return !(d.includes('DESCRIZIONI SERVIZI E VARIE') || d.includes('MARCA DA BOLLO'));
                    });
                if (!isArrivalSection && tasksToUse.length === 0) {
                    continue; // passa al prossimo gruppo, non mostrare header
                }

                // Per gli arrivi utilizza l'etichetta "OA" al posto di "OV" nel titolo del gruppo.
                const ovLabel = isArrivalSection ? 'OA' : 'OV';
                groupHeader.innerHTML = `${starIconHtml}${priorityIconHtml}${ovLabel}: ${group.ov} ${layoutIconHtml}`;
                // ========================================================
                // ==> FINE BLOCCO DI LOGICA UNIFICATO <==
                // ========================================================

                groupContainer.appendChild(groupHeader);

                tasksToUse.forEach(task => {
                    const taskElement = document.createElement('div');
                    taskElement.classList.add('gantt-task');
                    if (isArrivalSection) {
                        taskElement.classList.add(getFamilyColorClass(task.family));
                    } else {
                        taskElement.classList.add('shipping-task');
                        const codeStr = String(task.codiceArticolo || '').trim().toUpperCase();
                        if (codeStr.replace('*', '').startsWith('40')) {
                            taskElement.classList.add('medical-device-shipping', 'medical-device-shipping-priority');
                        } else if (codeStr.replace('*', '').startsWith('4') || ['7545', '40125V', '7316', '7317'].includes(codeStr)) {
                            taskElement.classList.add('medical-device-shipping');
                        }
                    }
                    const upperCaseCode = String(task.codiceArticolo || '').toUpperCase();
                    if (specialBorderCodes.has(upperCaseCode) || upperCaseCode.startsWith('PIL') || upperCaseCode.startsWith('EGC')) {
                        taskElement.classList.add('gantt-task-special-border');
                    }
                    taskElement.innerHTML = `<span class="task-code"><b>${task.codiceArticolo}</b></span><span class="task-details">${task.descrizioneArticolo}</span>`;

                    if (!isArrivalSection) {
                        // Sezione SPEDIZIONI: tooltip, lock, pallino CQ e flag QA
                        taskElement.addEventListener('mouseover', (e) => {
                            showSplitShippingTooltip(task, e);
                        });
                        taskElement.addEventListener('mouseout', hideGenericTooltip);

                        const lockIcon = document.createElement('span');
                        lockIcon.className = 'gantt-qa-lock';
                        lockIcon.innerHTML = '🔒';
                        lockIcon.dataset.rowId = task.rowId;
                        lockIcon.addEventListener('click', (e) => {
                            e.stopPropagation();
                            handleQACommentClick(e);
                        });
                        taskElement.appendChild(lockIcon);

                        // Pallino CQ (stato qualità) per le spedizioni
                        const cqDot = document.createElement('span');
                        cqDot.className = 'cq-status-dot';
                        cqDot.classList.add(`cq-status-${task.cqStatus || 'white'}`);
                        cqDot.dataset.rowId = task.rowId;
                        cqDot.addEventListener('click', (e) => {
                            e.stopPropagation();
                            handleCQStatusClick(e);
                        });
                        taskElement.appendChild(cqDot);

                        // Bandierina QA per le spedizioni
                        const qaFlag = document.createElement('span');
                        qaFlag.className = 'qa-status-flag';
                        qaFlag.classList.add(`qa-status-${task.qaStatus || 'white'}`);
                        qaFlag.dataset.rowId = task.rowId;
                        qaFlag.addEventListener('click', (e) => {
                            e.stopPropagation();
                            handleQAStatusClick(e);
                        });
                        taskElement.appendChild(qaFlag);

                        // Bandierina Spedizione per le spedizioni
                        const shipFlag = document.createElement('span');
                        shipFlag.className = 'ship-status-flag';
                        shipFlag.classList.add(`ship-status-${task.shipStatus || 'white'}`);
                        shipFlag.dataset.rowId = task.rowId;
                        shipFlag.textContent = 'S';
                        shipFlag.addEventListener('click', (e) => {
                            e.stopPropagation();
                            handleShipStatusClick(e);
                        });
                        taskElement.appendChild(shipFlag);

                        // Evidenzia le spedizioni ADR con una cornice rossa lampeggiante
                        const taskCodeUpper = (task.codiceArticolo || '').toString().trim().toUpperCase();
                        if (window.adrCodes && window.adrCodes.has(taskCodeUpper)) {
                            taskElement.classList.add('adr-shipping');
                        }
                    } else {
                        // Sezione ARRIVI: tooltip e pallino magazzino
                        taskElement.addEventListener('mouseover', (e) => showSplitShippingTooltip(task, e));
                        taskElement.addEventListener('mouseout', hideGenericTooltip);
                        // Pallino magazzino: indica se la merce è da evadere (bianco) o evasa (verde)
                        const magDot = document.createElement('span');
                        magDot.className = 'mag-status-dot';
                        magDot.classList.add(`mag-status-${task.magStatus || 'white'}`);
                        magDot.dataset.rowIndex = task.rowIndex;
                        magDot.addEventListener('click', (e) => {
                            e.stopPropagation();
                            handleMagStatusClick(e);
                        });
                        taskElement.appendChild(magDot);
                    }
                    groupContainer.appendChild(taskElement);
                });
                cell.appendChild(groupContainer);
            }
            ganttGrid.appendChild(cell);
        });
    };

    // Ottiene le righe di arrivo e costruisce un array di task che include
    // l'indice di riga (necessario per l'azione di evasione) e lo stato
    // magazzino (magStatus).  In questo modo il Gantt può filtrare le righe
    // già evase e consente di passare allo stato "evasa" direttamente dal Gantt.
    const arrivalRows = document.querySelectorAll('#arrivalScheduleTable tbody tr');
    const arrivalTasks = [];
    arrivalRows.forEach((r, idx) => {
        const rowData = getArrivalScheduleRowData(r);
        rowData.rowIndex = idx;
        arrivalTasks.push(rowData);
    });
    renderSectionRow('SPEDIZIONI IN USCITA', getAllShippingData(), 'shipping-group', false);
    renderSectionRow('MERCE IN ARRIVO', arrivalTasks, 'arrival-group', true);

    warehouseGanttChartDiv.appendChild(ganttGrid);

}
    function updateScrollButtons() {
        if (!tableContainer) return;
        const canScrollLeft = tableContainer.scrollLeft > 0;
        const canScrollRight = tableContainer.scrollLeft + tableContainer.clientWidth < tableContainer.scrollWidth;
        scrollLeftBtn.disabled = !canScrollLeft;
        scrollRightBtn.disabled = !canScrollRight;
    }

    scrollLeftBtn.addEventListener('click', () => tableContainer.scrollBy({ left: -200, behavior: 'smooth' }));
    scrollRightBtn.addEventListener('click', () => tableContainer.scrollBy({ left: 200, behavior: 'smooth' }));
    tableContainer.addEventListener('scroll', updateScrollButtons);
    window.addEventListener('resize', updateScrollButtons);

    productionTableBody.addEventListener('change', (event) => {
        if (event.target.matches('input, select')) {
            updateGanttChart();
            updateWarehouseGanttChart();
            updateDailyProductionTable();
            updateAnalisiTable();
            runFullCheck();
            autoSaveAllData();
        }
    });

    productionTableBody.addEventListener('input', (event) => {
        if (event.target.matches('.qty-requested-input, .stock-input, .qty-to-produce-input, .packaging-pieces-input, .packaging-kg-per-piece-input, .code-input, .production-days-input')) {
            validateRow(event.target.closest('tr'));
        }
        runFullCheck();
    });

    function removeHighlights() {
        document.querySelectorAll('.highlight').forEach(el => el.classList.remove('highlight'));
    }

    function findText() {
        removeHighlights();
        searchResults = [];
        currentSearchIndex = -1;
        const searchText = searchInput.value.trim().toLowerCase();
        if (!searchText) {
            findNextBtn.style.display = 'none';
            return;
        }

        productionTableBody.querySelectorAll('input, select').forEach(input => {
            if (input.value.toLowerCase().includes(searchText)) {
                searchResults.push(input);
            }
        });

        if (searchResults.length > 0) {
            findNext();
            findNextBtn.style.display = 'inline-block';
        } else {
            findNextBtn.style.display = 'none';
            showAlert('Nessun risultato trovato.');
        }
    }

    function findNext() {
        if (searchResults.length === 0) return;
        if (currentSearchIndex >= 0) searchResults[currentSearchIndex].classList.remove('highlight');
        currentSearchIndex = (currentSearchIndex + 1) % searchResults.length;
        const nextElement = searchResults[currentSearchIndex];
        nextElement.classList.add('highlight');
        nextElement.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
    }

    searchInput.addEventListener('input', () => {
        removeHighlights();
        searchResults = [];
        currentSearchIndex = -1;
        findNextBtn.style.display = 'none';
    });

    findBtn.addEventListener('click', findText);
    findNextBtn.addEventListener('click', findNext);

    function handleFilterColumnChange(selectElement, inputElement, flatpickrInstanceVar) {
        const column = selectElement.value;
        if (dateFilterColumns.includes(column)) {
            if (!flatpickrInstanceVar) {
                flatpickrInstanceVar = flatpickr(inputElement, {
                    dateFormat: "d/m/Y",
                    locale: "it",
                    allowInput: true
                });
            }
            inputElement.placeholder = "Seleziona data...";
        } else {
            if (flatpickrInstanceVar) {
                flatpickrInstanceVar.destroy();
                flatpickrInstanceVar = null;
            }
            inputElement.type = "text";
            inputElement.placeholder = "Valore filtro...";
        }
        if (selectElement.id === 'filterColumn1') {
            flatpickrInstance1 = flatpickrInstanceVar;
        } else if (selectElement.id === 'filterColumn2') {
            flatpickrInstance2 = flatpickrInstanceVar;
        }
    }

    filterColumn1Select.addEventListener('change', () => handleFilterColumnChange(filterColumn1Select, filterValue1Input, flatpickrInstance1));
    filterColumn2Select.addEventListener('change', () => handleFilterColumnChange(filterColumn2Select, filterValue2Input, flatpickrInstance2));

    // fix date filter
    function applyFilter() {
        const filterCol1 = filterColumn1Select.value;
        const filterVal1 = filterValue1Input.value.trim().toLowerCase();
        const filterCol2 = filterColumn2Select.value;
        const filterVal2 = filterValue2Input.value.trim().toLowerCase();

        productionTableBody.querySelectorAll('tr').forEach(row => {
            const rowData = getRowData(row);
            let showRow = true;

            if (filterCol1) {
                if (filterCol1 === medicalDevicesFilterValue) {
                    if (!isMedicalDeviceCode(rowData.codice)) {
                        showRow = false;
                    }
                } else if (dateFilterColumns.includes(filterCol1)) {
                    const filterDate = parseDateValue(filterVal1);
                    const rowDate = parseDateValue(rowData[filterCol1]);
                    if (filterDate && rowDate !== filterDate) {
                        showRow = false;
                    }
                } else if (!String(rowData[filterCol1] || '').toLowerCase().includes(filterVal1)) {
                    showRow = false;
                }
            }

            if (showRow && filterCol2) {
                if (filterCol2 === medicalDevicesFilterValue) {
                    if (!isMedicalDeviceCode(rowData.codice)) {
                        showRow = false;
                    }
                } else if (dateFilterColumns.includes(filterCol2)) {
                    const filterDate = parseDateValue(filterVal2);
                    const rowDate = parseDateValue(rowData[filterCol2]);
                    if (filterDate && rowDate !== filterDate) {
                        showRow = false;
                    }
                } else if (!String(rowData[filterCol2] || '').toLowerCase().includes(filterVal2)) {
                    showRow = false;
                }
            }
            row.style.display = showRow ? '' : 'none';
        });
        updateGanttChart();
        updateWarehouseGanttChart();
        updateDailyProductionTable();
        updateAnalisiTable();
        runFullCheck();
    }

    function clearFilter() {
        filterColumn1Select.value = '';
        filterValue1Input.value = '';
        if (flatpickrInstance1) {
            flatpickrInstance1.destroy();
            flatpickrInstance1 = null;
        }
        filterValue1Input.type = "text";
        filterValue1Input.placeholder = "Valore filtro 1...";

        filterColumn2Select.value = '';
        filterValue2Input.value = '';
        if (flatpickrInstance2) {
            flatpickrInstance2.destroy();
            flatpickrInstance2 = null;
        }
        filterValue2Input.type = "text";
        filterValue2Input.placeholder = "Valore filtro 2...";

        applyFilter();

        if (currentUserLevel === 1) {
            document.querySelectorAll('#searchInput, #filterColumn1, #filterValue1, #filterColumn2, #filterValue2, #applyFilterBtn, #clearFilterBtn').forEach(el => el.disabled = false);
        }
    }

    // VERSIONE AGGIORNATA
// Attiva i filtri in tempo reale
filterColumn1Select.addEventListener('change', applyFilter);
filterValue1Input.addEventListener('input', applyFilter);
filterColumn2Select.addEventListener('change', applyFilter);
filterValue2Input.addEventListener('input', applyFilter);
clearFilterBtn.addEventListener('click', clearFilter);
    clearFilterBtn.addEventListener('click', clearFilter);


    function updateStickyPositions() {
        if (stickyControlsWrapper && tableHead) {
            tableHead.style.top = `${stickyControlsWrapper.offsetHeight}px`;
        }
    }

    function getFormattedQuantityForDailyTable(rowData, type) {
        if (type === 'production') {
            return `${rowData.quantitaDaProdurre || ''} Kg`;
        }
        if (type === 'packaging') {
            const pezzi = rowData.confezionamentoPezzi || '';
            const kgPerPezzo = rowData.confezionamentoKgPerPiece || '';
            const unit = rowData.confezionamentoUnit || '';
            const isMedical = isMedicalDeviceCode(rowData.codice);

            if (isMedical) {
                return `${pezzi} ${kgPerPezzo}mL`;
            } else {
                return `${pezzi}X${kgPerPezzo}${unit}`;
            }
        }
        return '';
    }

    function getDailyProductionMachine(rowData, type) {
        if (type === 'packaging') {
            return "Confezionamento";
        }
        return rowData.macchinari || '';
    }

    function createDailyProductionRow(rowData = {}, rowType = 'production') {
        const row = document.createElement('tr');
        let className = '';
        const isMedical = isMedicalDeviceCode(rowType === 'production' ? rowData.codice : rowData.codiceConfezionamento);

        if (isMedical) {
            className = rowType === 'production' ? 'production-4xxxx-bg' : 'packaging-4xxxx-bg';
        } else {
            className = rowType === 'production' ? 'production-row-bg' : 'packaging-row-bg';
        }

        row.classList.add(className);

        row.innerHTML = `
            <td><input type="checkbox" class="daily-row-selector"></td>
            <td class="col-daily-op"><input type="text" value="${rowData.op || rowData.ope || ''}"></td>
            <td class="col-daily-ov"><input type="text" value="${rowData.ov || ''}"></td>
            <td class="col-daily-codice"><input type="text" value="${rowType === 'production' ? (rowData.codice || '') : (rowData.codiceConfezionamento || '')}"></td>
            <td class="col-daily-prodotto"><input type="text" value="${rowData.prodotto || ''}"></td>
            <td class="col-daily-cliente"><input type="text" value="${rowData.cliente || ''}"></td>
            <!-- Lotto viene spostato subito dopo il cliente -->
            <td class="col-daily-lotto"><input type="text" value="${rowData.lottoSC || rowData.lotto || ''}" style="text-align: left;"></td>
            <td class="col-daily-quantita"><input type="text" value="${rowType === 'production' ? getFormattedQuantityForDailyTable(rowData, 'production') : ''}"></td>
            <td class="col-daily-macchinario"><input type="text" value="${getDailyProductionMachine(rowData, rowType)}" list="macchinariOptionsListDaily"></td>
            <td class="col-daily-quantita-confez"><input type="text" value="${rowType === 'packaging' ? getFormattedQuantityForDailyTable(rowData, 'packaging') : ''}"></td>
            <td class="col-daily-operazioni">
                <select class="daily-operations-select">
                    ${dailyOperationsOptions.map(opt => `<option value="${opt}" ${rowData.operazioni === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                </select>
            </td>
            <td class="col-daily-operatori"><input type="text" value="${rowData.operatore || ''}" class="daily-operator-input"></td>
            <td class="col-daily-esito"><input type="text" value="${rowData.esito || ''}"></td>
            <td class="col-daily-qty-prodotta"><input type="text" value="${rowData.quantitaProdotta || ''}"></td>
            <!-- Lotto non viene più inserito in coda -->
            <td class="col-daily-data-avallo"><input type="text" value="${rowData.dataAvallo || ''}" style="width: 80px;"></td> `;

        const operationsSelect = row.querySelector('.daily-operations-select');
        operationsSelect.addEventListener('change', (e) => {
            if (e.target.value === 'Campo libero') {
                showPromptModal('Operazione Personalizzata', 'Descrizione:', '...').then(customText => {
                    if (customText) {
                        const newOption = new Option(customText, customText, false, true);
                        operationsSelect.add(newOption, operationsSelect.options[operationsSelect.options.length - 1]);
                        operationsSelect.value = customText;
                    } else {
                        operationsSelect.value = rowData.operazioni || '';
                    }
                });
            }
        });

        return row;
    }

    function getDailyRowData(row) {
        return {
            // Colonne OP e OV
            op: (row.querySelector('.col-daily-op input') || {}).value,
            ov: (row.querySelector('.col-daily-ov input') || {}).value,
            codice: row.querySelector('.col-daily-codice input').value,
            prodotto: row.querySelector('.col-daily-prodotto input').value,
            cliente: row.querySelector('.col-daily-cliente input').value,
            quantita: row.querySelector('.col-daily-quantita input').value,
            macchinario: row.querySelector('.col-daily-macchinario input').value,
            quantitaConfezionamento: row.querySelector('.col-daily-quantita-confez input').value,
            operazioni: row.querySelector('.daily-operations-select').value,
            operatore: row.querySelector('.col-daily-operatori input').value,
            esito: row.querySelector('.col-daily-esito input').value,
            quantitaProdotta: row.querySelector('.col-daily-qty-prodotta input').value,
            lotto: row.querySelector('.col-daily-lotto input').value,
            dataAvallo: row.querySelector('.col-daily-data-avallo input').value
        };
    }

/**
 * NUOVA FUNZIONE HELPER: Crea una chiave univoca e stabile per salvare/recuperare un commento.
 */
function getShippingRowStorageKey(rowData) {
    if (!rowData || !rowData.ov || !rowData.codiceArticolo) return null;
    // La chiave è basata sui dati dell'ordine, non su ID casuali
    return `qa_comment_${String(rowData.ov).trim()}_${String(rowData.codiceArticolo).trim()}`;
}


/**
 * MODIFICATA: Ora controlla il localStorage per recuperare i commenti salvati in modo persistente.
 */
/**
 * VERSIONE DEFINITIVA B - 1/2
 * Usa un ID stabile per garantire che il collegamento tra Gantt e tabella non si rompa mai.
 */
function createShippingScheduleRow(rowData = {}) {
    const row = document.createElement('tr');
    
    // [MODIFICA CHIAVE] Usa l'ID stabile invece di uno casuale
    row.dataset.rowId = getStableRowId(rowData);

    // Il resto della funzione rimane quasi identico...
    const storageKey = getShippingRowStorageKey(rowData);
    const savedComment = storageKey ? localStorage.getItem(storageKey) : null;
    rowData.commentiQA = savedComment || rowData.commentiQA || '';
    if (isMedicalDeviceCode(rowData.codiceArticolo)) {
        row.classList.add('production-4xxxx-bg');
    }
    const escapeAttr = (str) => String(str || '').replace(/"/g, '&quot;');
    row.innerHTML = `
        <td><input type="checkbox" class="shipping-row-selector"></td>
        <td><input type="text" value="${escapeAttr(rowData.ov || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.codiceArticolo || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.descrizioneArticolo || '')}" style="text-align: left;"></td>
        <td><input type="number" value="${escapeAttr(rowData.quantita || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.um || '')}"></td>
        <td><input type="text" class="datepicker" value="${escapeAttr(rowData.dataConsegna || '')}"></td>
        <td><input type="text" class="datepicker" value="${escapeAttr(rowData.dataConferma || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.ragioneSociale || '')}" style="text-align: left;"></td>
        <td><input type="text" value="${escapeAttr(rowData.riferimentoCliente || '')}" style="text-align: left;"></td>
        <td><input type="text" value="${escapeAttr(rowData.indirizzo || '')}" style="text-align: left;"></td>
        <td><input type="text" value="${escapeAttr(rowData.cap || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.citta || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.provincia || '')}"></td>
        <td><input type="text" value="${escapeAttr(rowData.telefono || '')}"></td>
    `;
    row.dataset.commentiQA = rowData.commentiQA;
    // Inizializza lo stato CQ per le spedizioni. Se non è presente nei dati,
    // viene impostato di default a 'white' (merce da analizzare).
    row.dataset.cqStatus = rowData.cqStatus || 'white';
    // Inizializza lo stato QA per le spedizioni. Se non è presente nei dati,
    // viene impostato di default a 'white' (merce in fase di valutazione).
    row.dataset.qaStatus = rowData.qaStatus || 'white';

    // Inizializza lo stato di spedizione. Se non è presente nei dati,
    // viene impostato di default a 'white' (ordine non ancora spedito).
    row.dataset.shipStatus = rowData.shipStatus || 'white';

    // Salva le note di servizio interne (colonna Q del file OS) nel dataset.  In questo modo
    // l'informazione sarà disponibile per il tooltip senza aggiungere una colonna visibile
    // nella tabella di spedizione.  Se non è presente alcuna nota, la stringa sarà vuota.
    row.dataset.noteServizio = rowData.noteServizio || '';
    row.querySelectorAll('.datepicker').forEach(input => {
        flatpickr(input, { dateFormat: "d/m/Y", locale: "it" });
    });
    row.querySelectorAll('input').forEach(input => {
        input.addEventListener('change', () => {
            updateWarehouseGanttChart();
            autoSaveAllData();
        });
    });
    return row;
}
/**
 * MODIFICATA: Garantisce che il campo commentiQA sia sempre letto dall'attributo data- della riga.
 */
function getShippingScheduleRowData(row) {
    const cells = row.cells;
    return {
        rowId: row.dataset.rowId,
        ov: cells[1].querySelector('input').value,
        codiceArticolo: cells[2].querySelector('input').value,
        descrizioneArticolo: cells[3].querySelector('input').value,
        quantita: cells[4].querySelector('input').value,
        um: cells[5].querySelector('input').value,
        dataConsegna: cells[6].querySelector('input').value,
        dataConferma: cells[7].querySelector('input').value,
        ragioneSociale: cells[8].querySelector('input').value,
        riferimentoCliente: cells[9].querySelector('input').value,
        indirizzo: cells[10].querySelector('input').value,
        cap: cells[11].querySelector('input').value,
        citta: cells[12].querySelector('input').value,
        provincia: cells[13].querySelector('input').value,
        telefono: cells[14].querySelector('input').value,
        commentiQA: row.dataset.commentiQA || '', // Legge sempre dall'attributo data aggiornato
        cqStatus: row.dataset.cqStatus || 'white', // Stato CQ dell'ordine di spedizione
        qaStatus: row.dataset.qaStatus || 'white', // Stato QA dell'ordine di spedizione
        // Stato Spedizione dell'ordine di spedizione
        shipStatus: row.dataset.shipStatus || 'white',
        // Restituisce anche le note interne (colonna Q) memorizzate nel dataset della riga.
        noteServizio: row.dataset.noteServizio || ''
    };
}

function getAllShippingData() {
    const data = [];
    document.querySelectorAll('#shippingScheduleTable tbody tr').forEach(row => {
        data.push(getShippingScheduleRowData(row));
    });
    return data;
}

    /**
     * Raccoglie i dati OPI memorizzati nel browser. Questi dati vengono inviati al server
     * assieme alle altre informazioni per renderli disponibili a tutti gli utenti.
     */
    function getOpiMonitorData() {
        const local = localStorage.getItem('opi_monitor_data');
        try {
            return local ? JSON.parse(local) : [];
        } catch (e) {
            return [];
        }
    }

    /**
     * Raccoglie i dati DeviceRef memorizzati nel browser.  Questi dati
     * contengono informazioni aggiuntive per i dispositivi/medicali (aghi,
     * siringhe, volumi, pesi, ecc.) importate tramite processDeviceRefFile().
     * Restituisce un array di oggetti o un array vuoto se non sono
     * presenti dati.
     */
    function getDeviceRefData() {
        const local = localStorage.getItem('deviceRefData');
        try {
            return local ? JSON.parse(local) : [];
        } catch (e) {
            return [];
        }
    }

// Rende le colonne di una tabella ridimensionabili
function makeTableResizable(table) {
    const headers = table.querySelectorAll('th');
    headers.forEach(header => {
        const resizer = document.createElement('div');
        resizer.className = 'resizer';
        header.appendChild(resizer);

        resizer.addEventListener('mousedown', (e) => {
            e.preventDefault();
            const startX = e.pageX;
            const startWidth = header.offsetWidth;

            const onMouseMove = (e) => {
                const newWidth = startWidth + (e.pageX - startX);
                if (newWidth > 30) { // Larghezza minima
                    header.style.width = `${newWidth}px`;
                }
            };

            const onMouseUp = () => {
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
                document.body.classList.remove('resizing');
            };

            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
            document.body.classList.add('resizing');
        });
    });
}

// ===================================================================
// ==> GESTIONE TIMESTAMP ULTIMI IMPORT <==
// ===================================================================
/**
 * Formatta una data in formato "dd/mm/aaaa HH:MM" per visualizzare
 * l'orario degli import in modo compatto (senza secondi).
 * @param {Date} date
 * @returns {string}
 */
function formatDateTimeForDisplay(date) {
    return date.toLocaleDateString('it-IT') + ' ' + date.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
}

/**
 * Aggiorna tutte le etichette che mostrano l'ultima data/ora di importazione.
 * Legge i valori dal localStorage e aggiorna il testo degli elementi
 * con id corrispondenti (es. lastImportPP, lastImportOV, etc.).
 */
function updateImportTimestamps() {
    const mapping = {
        'PP': 'lastImportPP',
        'OV': 'lastImportOV',
        'OPI': 'lastImportOPI',
        'OS': 'lastImportOS',
        'Arrivals': 'lastImportArrivals',
        'Layout': 'lastImportLayout',
        'referenze': 'lastImportReferenze',
        'pianoAnalitico': 'lastImportPianoAnalitico',
        'deviceRef': 'lastImportDeviceRef',
        'medicalProduction': 'lastImportMedicalProduction',
        // Importazione inventario per la tabella Merce in Scadenza
        'inventory': 'lastImportInventory'
    };
    Object.keys(mapping).forEach(mode => {
        const spanId = mapping[mode];
        const spanElem = document.getElementById(spanId);
        if (spanElem) {
        // Recupera il valore direttamente in base all'ID dello span
        const ts = localStorage.getItem(spanId);
        spanElem.textContent = ts ? ` (Ultimo import: ${ts})` : '';
        }
    });

    // Aggiorna anche la sezione riassuntiva globale con tutte le date/ore.
    updateLastImportsSummary();
}

/**
 * Aggiorna automaticamente le colonne OPE (Ordine di Produzione Esterno) e OV
 * nella tabella giornaliera.  Per ogni riga della tabella giornaliera, se il
 * codice prodotto corrisponde a un record nella tabella OPI, copia l'OP
 * (salvato come OPE), l'OV, l'operatore e il lotto nelle colonne
 * corrispondenti qualora siano vuote.  Questa funzione non sovrascrive
 * valori già presenti inseriti manualmente dall'utente.
 */
function updateDailyOpeOv() {
    try {
        const opiRows = Array.from(document.querySelectorAll('#opiTable tbody tr'));
        const opiMap = {};
        opiRows.forEach(row => {
            const cells = row.querySelectorAll('td');
            const op = cells[1] ? cells[1].textContent.trim() : '';
            const ov = cells[2] ? cells[2].textContent.trim() : '';
            const codice = cells[3] ? cells[3].textContent.trim().toUpperCase() : '';
            const operatore = cells[9] ? cells[9].textContent.trim() : '';
            const lotto = cells[6] ? cells[6].textContent.trim() : '';
            if (!opiMap[codice]) opiMap[codice] = [];
            opiMap[codice].push({ op, ov, operatore, lotto });
        });
        document.querySelectorAll('#dailyProductionTable tbody tr').forEach(row => {
            const codiceInput = row.querySelector('.col-daily-codice input');
            if (!codiceInput) return;
            const codeVal = (codiceInput.value || '').trim().toUpperCase();
            const matches = opiMap[codeVal];
            if (matches && matches.length > 0) {
                const { op, ov, operatore, lotto } = matches[0];
                const opInput = row.querySelector('.col-daily-op input');
                const ovInput = row.querySelector('.col-daily-ov input');
                const operatorInput = row.querySelector('.col-daily-operatori input');
                const lottoInput = row.querySelector('.col-daily-lotto input');
                if (opInput && !opInput.value) opInput.value = op || '';
                if (ovInput && !ovInput.value) ovInput.value = ov || '';
                if (operatorInput && !operatorInput.value) operatorInput.value = operatore || '';
                if (lottoInput && !lottoInput.value) lottoInput.value = lotto || '';
            }
        });
    } catch (e) {
        console.warn('Errore in updateDailyOpeOv:', e);
    }
}

/**
 * Aggiorna dinamicamente il datalist utilizzato per suggerire i nomi degli
 * operatori nel programma giornaliero.  Il datalist viene popolato
 * analizzando le righe correnti della tabella giornaliera e aggiungendo
 * ciascun nome operatore unico.  Se nessuna riga è visibile o non sono
 * presenti operatori, il datalist viene svuotato.  Eventuali errori
 * vengono registrati ma non interrompono l'esecuzione del programma.
 */
function updateOperatorSuggestions() {
    try {
        if (!operatorSuggestionsList) return;
        const ops = new Set();
        // Itera sulle righe visibili della tabella e raccoglie gli operatori
        document.querySelectorAll('#dailyProductionTable tbody tr').forEach(row => {
            if (row.style.display === 'none') return; // ignora righe filtrate
            const input = row.querySelector('.col-daily-operatori input');
            const val = input ? input.value.trim() : '';
            if (val) ops.add(val);
        });
        let optionsHtml = '';
        ops.forEach(op => {
            const sanitized = op.replace(/"/g, '&quot;');
            optionsHtml += `<option value="${sanitized}"></option>`;
        });
        operatorSuggestionsList.innerHTML = optionsHtml;
    } catch (e) {
        console.warn('Errore nell\'aggiornamento del datalist operatori:', e);
    }
}

/**
 * Aggiorna il datalist dei macchinari per la tabella giornaliera.  Oltre
 * all'elenco predefinito contenuto in macchinariOptions, vengono
 * aggiunti eventuali macchinari inseriti manualmente nelle righe
 * esistenti.  In questo modo l'utente può selezionare rapidamente un
 * macchinario già utilizzato o scegliere uno dei valori suggeriti.
 */
function updateMacchinariOptionsListDaily() {
    try {
        if (!macchinariOptionsListDaily) return;
        const machineSet = new Set(macchinariOptions);
        // Aggiungi i macchinari presenti nelle righe correnti
        document.querySelectorAll('#dailyProductionTable tbody tr').forEach(row => {
            const input = row.querySelector('.col-daily-macchinario input');
            const val = input ? input.value.trim() : '';
            if (val) machineSet.add(val);
        });
        let optionsHtml = '';
        // Ordina alfabeticamente per presentazione coerente
        Array.from(machineSet).sort((a,b) => a.localeCompare(b)).forEach(mach => {
            const sanitized = mach.replace(/"/g, '&quot;');
            optionsHtml += `<option value="${sanitized}"></option>`;
        });
        macchinariOptionsListDaily.innerHTML = optionsHtml;
    } catch (e) {
        console.warn('Errore nell\'aggiornamento del datalist macchinari giornaliero:', e);
    }
}

/**
 * Ordina le righe di una tabella in base al valore della colonna specificata.
 * Riconosce automaticamente numeri, date (formato gg/mm/aaaa) e stringhe.
 * L'ordinamento avviene in loco e l'ordine viene invertito ad ogni click.
 * @param {HTMLTableElement} table La tabella da ordinare
 * @param {number} colIndex Indice della colonna su cui ordinare
 * @param {boolean} ascending True per ordine crescente, false per decrescente
 */
function sortTableByColumn(table, colIndex, ascending) {
    const tbody = table.tBodies[0];
    if (!tbody) return;
    const rowsArray = Array.from(tbody.querySelectorAll('tr'));
    // Funzione di supporto: estrae il valore da una cella. Se contiene un input o un select,
    // restituisce il valore dell'input o l'opzione selezionata; altrimenti restituisce
    // il testo contenuto nella cella.
    const getCellValue = (cell) => {
        if (!cell) return '';
        const inputEl = cell.querySelector('input, select');
        if (inputEl) {
            if (inputEl.tagName.toLowerCase() === 'select') {
                return (inputEl.options[inputEl.selectedIndex] ? inputEl.options[inputEl.selectedIndex].text : '').trim();
            } else {
                return (inputEl.value || '').trim();
            }
        }
        return cell.textContent.trim();
    };
    rowsArray.sort((a, b) => {
        const cellA = a.cells[colIndex];
        const cellB = b.cells[colIndex];
        const valA = getCellValue(cellA);
        const valB = getCellValue(cellB);
        // Prova a interpretare come data nel formato italiano (dd/mm/yy o dd/mm/yyyy)
        const datePattern = /^\d{1,2}\/\d{1,2}\/\d{2,4}/;
        let comparison = 0;
        if (datePattern.test(valA) && datePattern.test(valB)) {
            const parseItDate = (str) => {
                const parts = str.split('/');
                const d = parseInt(parts[0], 10);
                const m = parseInt(parts[1], 10);
                let y = parseInt(parts[2], 10);
                if (parts[2].length === 2) {
                    y = y > 50 ? (1900 + y) : (2000 + y);
                }
                return new Date(y, m - 1, d);
            };
            const dateA = parseItDate(valA);
            const dateB = parseItDate(valB);
            comparison = dateA - dateB;
        } else {
            // Prova a interpretare come numero (ignorando separatori)
            const numA = parseFloat(valA.replace(/\./g, '').replace(',', '.'));
            const numB = parseFloat(valB.replace(/\./g, '').replace(',', '.'));
            if (!isNaN(numA) && !isNaN(numB)) {
                comparison = numA - numB;
            } else {
                comparison = valA.localeCompare(valB);
            }
        }
        return ascending ? comparison : -comparison;
    });
    rowsArray.forEach(row => tbody.appendChild(row));
}

/**
 * Rende una tabella ordinabile aggiungendo un'icona di ordinamento ad ogni intestazione
 * e applicando la logica di sortTableByColumn al click. La prima colonna (checkbox) viene
 * ignorata. L'icona mostra una freccia "↕" per indicare la possibilità di ordinamento.
 * @param {HTMLTableElement} table La tabella da rendere ordinabile
 */
function makeTableSortable(table) {
    if (!table || !table.tHead || !table.tBodies.length) return;
    const ths = table.tHead.querySelectorAll('th');
    ths.forEach((th, index) => {
        // Salta la prima colonna (checkbox) o colonne vuote
        if (index === 0 || th.classList.contains('no-sort')) return;
        // Aggiungi icona di ordinamento solo se non già presente
        if (!th.querySelector('.sort-icon')) {
            const icon = document.createElement('span');
            icon.className = 'sort-icon';
            icon.textContent = '↕';
            th.appendChild(icon);
        }
        let ascending = true;
        th.style.cursor = 'pointer';
        th.addEventListener('click', (e) => {
            // Evita l'ordinamento se si clicca su un elemento interattivo all'interno dell'intestazione
            const tag = e.target.tagName.toLowerCase();
            if (tag === 'input' || tag === 'select' || tag === 'button' || tag === 'svg' || tag === 'path') return;
            sortTableByColumn(table, index, ascending);
            ascending = !ascending;
        });
    });
}

// ---------------------------------------------------------------------
// Fallback definitions per le funzioni makeTableResizable e makeTableSortable.
// In alcune versioni del file di base queste funzioni potrebbero non essere
// definite quando sono invocate in altri contesti.  Qui garantiamo che
// esistano sempre come no-op, così da evitare errori 'ReferenceError'.
if (typeof window !== 'undefined') {
  if (typeof window.makeTableResizable !== 'function') {
    window.makeTableResizable = function() {};
  }
  if (typeof window.makeTableSortable !== 'function') {
    window.makeTableSortable = function() {};
  }
  // Crea una definizione di fallback per addLogEntry nel caso in cui non sia
  // stata definita a livello globale. Questo evita errori ReferenceError
  // quando alcune parti del codice tentano di chiamarla. La funzione no-op
  // accetta qualsiasi argomento ma non esegue alcuna azione.
  if (typeof window.addLogEntry !== 'function') {
    window.addLogEntry = function() {
      // Se non è disponibile una definizione reale, usa console.log per debug.
      try {
        if (typeof console !== 'undefined' && typeof console.log === 'function') {
          console.log.apply(console, arguments);
        }
      } catch (e) {
        // ignora eventuali errori della console
      }
    };
  }
}

/*
 * Rende trascinabile un pop-up di allerta (ADR o CQ/QA).  L'utente può
 * spostare la finestra facendo clic e trascinando su qualsiasi area del
 * pop-up che non sia un pulsante o un'icona di chiusura.  Durante il
 * trascinamento vengono aggiornate le proprietà CSS 'left' e 'top' per
 * posizionare il pop-up.  La trasformazione iniziale (transform)
 * viene rimossa quando l'utente inizia a trascinare.
 */
function makeAlertDraggable(alertId) {
    const alertEl = document.getElementById(alertId);
    if (!alertEl) return;
    let dragging = false;
    let offsetX = 0;
    let offsetY = 0;
    const onMouseDown = (e) => {
        // Solo tasto sinistro
        if (e.button !== 0) return;
        const target = e.target;
        /*
         * Consenti l'avvio del drag ovunque all'interno della finestra di allerta
         * tranne che sui pulsanti, sulle icone di chiusura e nella barra dei
         * pulsanti.  Questo evita che un click sull'elenco o sulla cornice
         * impedisca il trascinamento.  Usiamo closest() per verificare se
         * l'elemento (o uno dei suoi genitori) è all'interno del container
         * dedicato ai pulsanti.
         */
        if (target.closest('button') ||
            target.closest('.adr-alert-buttons') ||
            target.closest('.quality-alert-buttons') ||
            target.closest('.warehouse-alert-buttons') ||
            target.classList.contains('adr-close-btn') ||
            target.classList.contains('quality-close-btn') ||
            target.classList.contains('warehouse-close-btn')) {
            return;
        }
        dragging = true;
        const rect = alertEl.getBoundingClientRect();
        offsetX = e.clientX - rect.left;
        offsetY = e.clientY - rect.top;
        // Rimuovi la trasformazione predefinita per consentire lo spostamento tramite left/top
        alertEl.style.transform = '';
        // Assicura che la posizione sia fissata per gli spostamenti successivi
        if (getComputedStyle(alertEl).position !== 'fixed') {
            alertEl.style.position = 'fixed';
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
        e.preventDefault();
    };
    const onMouseMove = (e) => {
        if (!dragging) return;
        alertEl.style.left = (e.clientX - offsetX) + 'px';
        alertEl.style.top = (e.clientY - offsetY) + 'px';
    };
    const onMouseUp = () => {
        dragging = false;
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
    };
    alertEl.addEventListener('mousedown', onMouseDown);
}

// Una volta che il DOM è pronto, rendi tutte le tabelle ordinabili
document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('table').forEach(tbl => {
        makeTableSortable(tbl);
    });
    // Rendi trascinabili i pop-up ADR, Quality e Magazzino se esistono
    makeAlertDraggable('adrNotification');
    makeAlertDraggable('qualityNotification');
    makeAlertDraggable('warehouseNotification');

    // Aggiorna subito il Gantt di magazzino per inizializzare l'intervallo a 30 giorni
    if (typeof updateWarehouseGanttChart === 'function') {
        try {
            updateWarehouseGanttChart();
            // Dopo l'aggiornamento iniziale del Gantt di magazzino, crea le barre di scorrimento in linea
            if (typeof addInlineScrollbarsToWarehouseGantt === 'function') {
                addInlineScrollbarsToWarehouseGantt();
            }
            // Inizializza anche la barra di scorrimento esterna e i pulsanti laterali
            if (typeof initWarehouseGanttExternalScrollbar === 'function') {
                initWarehouseGanttExternalScrollbar();
            }
        } catch (err) {
            console.warn('Errore durante la generazione iniziale del Gantt magazzino:', err);
        }
    }

        // Attiva i pulsanti di scorrimento per il Gantt di magazzino.  Questi
        // pulsanti consentono di spostare orizzontalmente il contenitore
        // senza usare la barra di scorrimento.  Scrolliamo di alcune colonne
        // Rimozione dei pulsanti di scorrimento per il Gantt di magazzino.
        // Si utilizza la barra di scorrimento nativa del contenitore per lo
        // spostamento orizzontale.  Pertanto non registriamo alcun listener
        // per i pulsanti e non riposizioniamo elementi.  La logica qui
        // presente è stata rimossa.

        /* -------------------------------------------------------------------------
         * Inizializzazione delle barre di scorrimento "fantasma" per le tabelle
         * giornaliere di spedizione e arrivo merce.  Queste barre consentono
         * all'utente di scorrere orizzontalmente quando la barra originale
         * non è visibile a causa dello scorrimento verticale.  Vengono
         * create al primo caricamento e sincronizzate con il rispettivo
         * contenitore di tabella.
         * ----------------------------------------------------------------------- */
        try {
            const arrivalWrapper = document.querySelector('#arrivalScheduleContainer .daily-production-table-wrapper');
            const shippingWrapper = document.querySelector('#shippingScheduleContainer .daily-production-table-wrapper');
            if (arrivalWrapper) {
                createDockedScrollbarsForContainer(arrivalWrapper);
            }
            if (shippingWrapper) {
                createDockedScrollbarsForContainer(shippingWrapper);
            }
            // Non aggiungiamo barre ancorate al Gantt Magazzino: utilizziamo la barra nativa del contenitore

    // Ridefinisci updateWarehouseGanttChart per assicurare che le barre di scorrimento in linea vengano
    // rigenerate dopo ogni aggiornamento del Gantt.  Questo mantiene la barra orizzontale adiacente al grafico.
    if (typeof updateWarehouseGanttChart === 'function' && typeof addInlineScrollbarsToWarehouseGantt === 'function') {
        const _origWarehouseGantt = updateWarehouseGanttChart;
        updateWarehouseGanttChart = function() {
            _origWarehouseGantt.apply(this, arguments);
            addInlineScrollbarsToWarehouseGantt();
            // Dopo aver rigenerato le barre interne, inizializza la barra esterna e i pulsanti laterali
            if (typeof initWarehouseGanttExternalScrollbar === 'function') {
                initWarehouseGanttExternalScrollbar();
            }
        };
    }
        } catch (e) {
            console.warn('Impossibile inizializzare le barre di scorrimento fantasma:', e);
        }
});

/*
 * Aggiorna la sezione riassuntiva che mostra l'ultimo import per ciascun tipo di file.
 * Viene chiamata da updateImportTimestamps() per mantenere sincronizzata la sezione.
 */
function updateLastImportsSummary() {
    const summaryDiv = document.getElementById('lastImportsSummary');
    if (!summaryDiv) return;
    const displayNames = {
        PP: 'PP',
        OV: 'OV',
        OPI: 'OPI',
        OS: 'OS',
        Arrivals: 'Arrivi',
        Layout: 'Layout',
        referenze: 'Referenze',
        pianoAnalitico: 'Piano Analitico',
        deviceRef: 'DeviceRef',
        medicalProduction: 'Prod. MD'
    };
    let html = '';
    Object.keys(displayNames).forEach(mode => {
        // Usa chiavi uniformi senza underscore (es. lastImportPP, lastImportOV)
        // Normalizza il nome per costruire la chiave senza underscore e con iniziale maiuscola
        const normalized = mode.charAt(0).toUpperCase() + mode.slice(1);
        const ts = localStorage.getItem('lastImport' + normalized);
        const label = displayNames[mode];
        html += `<div><strong>Import ${label}:</strong> ${ts ? ts : '—'}</div>`;
    });
    summaryDiv.innerHTML = html;
}

/*
 * Inizializza la barra di scorrimento esterna e i pulsanti laterali per il
 * Gantt di magazzino.  Questa funzione crea un secondo controllo
 * orizzontale tra la tabella delle spedizioni giornaliere e il grafico
 * Gantt, consentendo all'utente di scorrere orizzontalmente il Gantt
 * senza dover raggiungere l'estremità inferiore.  Inoltre sincronizza
 * l'input tra la barra esterna e il contenitore scrollabile del Gantt
 * e abilita/disabilita i pulsanti di scorrimento laterali in base alla
 * posizione corrente.  La funzione è idempotente: se è già stata
 * inizializzata, effettua un refresh delle misure e ritorna.
 */
function initWarehouseGanttExternalScrollbar() {
    const container = document.getElementById('warehouseGanttChartContainer');
    const wrapper = document.getElementById('warehouseGanttScrollWrapper');
    const externalBar = document.getElementById('warehouseGanttExternalScrollbar');
    const sizer = externalBar ? externalBar.querySelector('.gantt-external-sizer') : null;
    const btnWrapper = document.getElementById('warehouseGanttScrollButtonsWrapper');
    const btnLeft = document.getElementById('warehouseGanttScrollLeftBtn');
    const btnRight = document.getElementById('warehouseGanttScrollRightBtn');
    // Se uno degli elementi richiesti non esiste, abbandona.
    if (!container || !wrapper || !externalBar || !sizer || !btnWrapper || !btnLeft || !btnRight) return;

    // Funzione che aggiorna la larghezza del sizer in base alla larghezza del Gantt.
    function refreshSizerWidth() {
        try {
            const w = Math.max(wrapper.scrollWidth || 0, container.clientWidth || 0);
            sizer.style.width = w + 'px';
        } catch (e) {}
    }

    // Funzione che abilita o disabilita i pulsanti laterali in base alla posizione.
    function updateButtonState() {
        const maxScroll = (wrapper.scrollWidth || 0) - (container.clientWidth || 0);
        // Disabilita il pulsante sinistro se siamo all'inizio
        if (wrapper.scrollLeft <= 0) {
            btnLeft.setAttribute('disabled', 'true');
        } else {
            btnLeft.removeAttribute('disabled');
        }
        // Disabilita il pulsante destro se siamo alla fine
        if (wrapper.scrollLeft >= maxScroll - 1) {
            btnRight.setAttribute('disabled', 'true');
        } else {
            btnRight.removeAttribute('disabled');
        }
    }

    // Funzione che riposiziona verticalmente il wrapper dei pulsanti al centro
    // della porzione visibile del Gantt.  In questo modo i pulsanti restano
    // accessibili anche quando la pagina viene scrollata.
    function repositionButtons() {
        const rect = container.getBoundingClientRect();
        // Calcola la parte visibile del Gantt nel viewport
        const visibleTop = Math.max(rect.top, 0);
        const visibleBottom = Math.min(rect.bottom, window.innerHeight);
        const visibleHeight = Math.max(visibleBottom - visibleTop, 0);
        const wrapperHeight = btnWrapper.offsetHeight || 0;
        const top = visibleTop + (visibleHeight / 2) - (wrapperHeight / 2);
        // Se fuori viewport, posiziona comunque in cima
        // Il wrapper dei pulsanti è posizionato in modo fisso (fixed), quindi
        // la coordinata top è relativa al viewport.  Non aggiungiamo
        // window.scrollY.
        btnWrapper.style.top = top + 'px';
    }

// --- PATCH: keep Warehouse Gantt buttons aligned on scroll/resize ---
try {
  if (typeof repositionButtons === 'function') {
    window.addEventListener('scroll', repositionButtons, { passive: true });
    window.addEventListener('resize', function() {
      try { if (typeof refreshSizerWidth === 'function') { refreshSizerWidth(); } } catch (e) {}
      repositionButtons();
    });
  }
} catch (e) {
  console.warn('Patch: could not bind scroll/resize for Warehouse Gantt buttons', e);
}
// --- END PATCH ---


    // Se già inizializzato, aggiorna larghezza e posizione e ritorna
    if (externalBar.__initialized) {
        refreshSizerWidth();
        updateButtonState();
        repositionButtons();
        return;
    }
    externalBar.__initialized = true;

    // Aggiorna larghezza iniziale del sizer e stato dei pulsanti
    refreshSizerWidth();
    updateButtonState();
    repositionButtons();

    // Sincronizza lo scroll tra la barra esterna e il wrapper del Gantt
    let syncingFromBar = false;
    let syncingFromWrapper = false;
    wrapper.addEventListener('scroll', () => {
        if (syncingFromBar) return;
        syncingFromWrapper = true;
        externalBar.scrollLeft = wrapper.scrollLeft;
        syncingFromWrapper = false;
        updateButtonState();
    }, { passive: true });
    externalBar.addEventListener('scroll', () => {
        if (syncingFromWrapper) return;
        syncingFromBar = true;
        wrapper.scrollLeft = externalBar.scrollLeft;
        syncingFromBar = false;
        updateButtonState();
    }, { passive: true });

    // Gestisce i click sui pulsanti laterali per spostare lo scroll del Gantt
    const scrollStep = 330;
    btnLeft.addEventListener('click', () => {
        wrapper.scrollLeft = Math.max(wrapper.scrollLeft - scrollStep, 0);
    });
    btnRight.addEventListener('click', () => {
        const maxScroll = (wrapper.scrollWidth || 0) - (container.clientWidth || 0);
        wrapper.scrollLeft = Math.min(wrapper.scrollLeft + scrollStep, maxScroll);
    });

    // Aggiorna la larghezza del sizer e la posizione dei pulsanti su resize
    window.addEventListener('resize', () => {
        refreshSizerWidth();
        updateButtonState();
        repositionButtons();
    });
    // Aggiorna la posizione dei pulsanti su scroll della finestra
    window.addEventListener('scroll', repositionButtons, { passive: true });

    // Osserva il wrapper e il contenitore per variazioni di dimensione
    try {
        const ro = new ResizeObserver(() => {
            refreshSizerWidth();
            updateButtonState();
        });
        ro.observe(wrapper);
        ro.observe(container);
        externalBar._ro = ro;
    } catch (e) {}
}

/* ===========================================================================
 * Gestione delle barre di scorrimento orizzontale "fantasma" per le tabelle.
 * Questa funzione crea una barra di scorrimento indipendente che rimane
 * visibile a metà altezza della tabella corrente.  La barra viene
 * sincronizzata con la barra di scorrimento nativa del contenitore
 * (tipicamente una .daily-production-table-wrapper).  L'altezza e la
 * larghezza vengono calcolate dinamicamente in base alla porzione visibile.
 */

/* ============================================================
 * Barre orizzontali ANCORATE sopra/sotto ogni contenitore
 * (tabelle giornaliere e Gantt magazzino), sincronizzate tra loro
 * e con la barra nativa del contenitore.
 * ============================================================ */
function createDockedScrollbarsForContainer(container) {
    if (!container || container.__hasDockScrollbars) return;
    container.__hasDockScrollbars = true;
    container.classList.add('has-dock-scrollbars');

    // Crea le due barre (sopra e sotto) con l'inner che determina la larghezza
    const topBar = document.createElement('div');
    topBar.className = 'dock-scrollbar top';
    const topInner = document.createElement('div');
    topInner.style.height = '1px';
    topBar.appendChild(topInner);

    const bottomBar = document.createElement('div');
    bottomBar.className = 'dock-scrollbar bottom';
    const bottomInner = document.createElement('div');
    bottomInner.style.height = '1px';
    bottomBar.appendChild(bottomInner);

    // Inserisci la barra sopra come primo figlio e la barra sotto come ultimo
    container.insertBefore(topBar, container.firstChild);
    container.appendChild(bottomBar);

    function syncSizes() {
        // Larghezza massima scrollabile del contenitore
        const width = container.scrollWidth;
        topInner.style.width = width + 'px';
        bottomInner.style.width = width + 'px';
        // Allinea la posizione delle barre allo scroll corrente
        topBar.scrollLeft = container.scrollLeft;
        bottomBar.scrollLeft = container.scrollLeft;
    }

    // Sincronizzazione bidirezionale
    topBar.addEventListener('scroll', () => {
        if (container.scrollLeft !== topBar.scrollLeft) {
            container.scrollLeft = topBar.scrollLeft;
        }
        if (bottomBar.scrollLeft !== topBar.scrollLeft) {
            bottomBar.scrollLeft = topBar.scrollLeft;
        }
    });

    bottomBar.addEventListener('scroll', () => {
        if (container.scrollLeft !== bottomBar.scrollLeft) {
            container.scrollLeft = bottomBar.scrollLeft;
        }
        if (topBar.scrollLeft !== bottomBar.scrollLeft) {
            topBar.scrollLeft = bottomBar.scrollLeft;
        }
    });

    container.addEventListener('scroll', () => {
        if (topBar.scrollLeft !== container.scrollLeft) {
            topBar.scrollLeft = container.scrollLeft;
        }
        if (bottomBar.scrollLeft !== container.scrollLeft) {
            bottomBar.scrollLeft = container.scrollLeft;
        }
    });

    // Osserva cambi dimensioni per adeguare la larghezza dell'inner
    if (typeof ResizeObserver !== 'undefined') {
        const ro = new ResizeObserver(syncSizes);
        ro.observe(container);
        // osserva anche il primo figlio "reale" se presente (tabella o chart)
        const firstRealChild = Array.from(container.children).find(el => !el.classList.contains('dock-scrollbar'));
        if (firstRealChild) ro.observe(firstRealChild);
    } else {
        window.addEventListener('resize', syncSizes);
        setInterval(syncSizes, 500);
    }

    // Impostazione iniziale
    setTimeout(syncSizes, 0);
}
function createGhostScrollbarForContainer(container) {
    if (!container) return;
    // Crea la barra fantasma e il suo contenitore interno per impostare la larghezza
    const ghost = document.createElement('div');
    ghost.className = 'ghost-scrollbar';
    const inner = document.createElement('div');
    inner.style.height = '1px';
    ghost.appendChild(inner);
    document.body.appendChild(ghost);

    // Aggiorna la larghezza interna in base all'area scrollabile del contenitore
    function updateInnerWidth() {
        inner.style.width = container.scrollWidth + 'px';
    }

/* Barre di scorrimento ancorate al Gantt Magazzino (sopra e sotto la griglia) */
function addInlineScrollbarsToWarehouseGantt() {

  // Create (if needed) a native bottom scroller wrapper that holds the chart.
  const container = document.getElementById('warehouseGanttChartContainer');
  const chart = document.getElementById('warehouseGanttChart');
  if (!container || !chart) return;

  let wrapper = document.getElementById('warehouseGanttScrollWrapper');
  if (!wrapper) {
    wrapper = document.createElement('div');
    wrapper.id = 'warehouseGanttScrollWrapper';
    wrapper.className = 'gantt-inline-wrapper';
    // Insert wrapper before chart and move chart inside
    chart.parentNode.insertBefore(wrapper, chart);
    wrapper.appendChild(chart);
  }
  // Ensure container itself does not show its own horizontal scrollbar
  try {
    container.style.overflowX = 'hidden';
    container.style.overflowY = 'visible';
  } catch(e){}

  // === TOP INLINE SCROLLBAR (synced) ===
  let topBar = document.getElementById('warehouseGanttTopScrollbar');
  if (!topBar) {
    topBar = document.createElement('div');
    topBar.id = 'warehouseGanttTopScrollbar';
    topBar.className = 'gantt-inline-scrollbar';
    // add a sizer to mirror content width
    const s = document.createElement('div');
    s.className = 'gantt-inline-sizer';
    s.style.height = '1px';
    topBar.appendChild(s);
    // insert the top bar just above the wrapper so it appears "sopra il Gantt"
    container.insertBefore(topBar, wrapper);
  }
  const sizer = topBar.querySelector('.gantt-inline-sizer') || topBar.firstElementChild;

  function refreshSizerWidth() {
    try {
      // mirror the chart scrollWidth; ensure at least container width
      const w = Math.max(chart.scrollWidth || 0, container.clientWidth || 0);
      sizer.style.width = w + 'px';
    } catch(e){}
  }
  refreshSizerWidth();

  // sync scroll positions between top bar and wrapper
  if (!topBar._syncBound) {
    let syncingFromTop = false;
    let syncingFromWrapper = false;
    wrapper.addEventListener('scroll', function() {
      if (syncingFromTop) return;
      syncingFromWrapper = true;
      topBar.scrollLeft = wrapper.scrollLeft;
      syncingFromWrapper = false;
    }, { passive: true });

    topBar.addEventListener('scroll', function() {
      if (syncingFromWrapper) return;
      syncingFromTop = true;
      wrapper.scrollLeft = topBar.scrollLeft;
      syncingFromTop = false;
    }, { passive: true });

    // initialize alignment
    topBar.scrollLeft = wrapper.scrollLeft;

    // observers to keep width in sync
    try {
      const ro = new ResizeObserver(refreshSizerWidth);
      ro.observe(chart);
      ro.observe(wrapper);
      ro.observe(container);
      topBar._ro = ro;
    } catch(e){}
    try {
      const mo = new MutationObserver(refreshSizerWidth);
      mo.observe(chart, { childList: true, subtree: true, attributes: true });
      topBar._mo = mo;
    } catch(e){}

    window.addEventListener('resize', refreshSizerWidth);
    topBar._syncBound = true;
  }

}



    // Reposiziona la barra fantasma in base alla porzione visibile del contenitore
    function repositionGhost() {
        const rect = container.getBoundingClientRect();
        // Se la tabella è fuori dal viewport, nascondi la barra
        if (rect.bottom < 0 || rect.top > window.innerHeight) {
            ghost.style.display = 'none';
            return;
        }
        // Mostra la barra e imposta larghezza e posizione orizzontale
        ghost.style.display = 'block';
        ghost.style.width = rect.width + 'px';
        ghost.style.left = rect.left + 'px';
        // Calcola la porzione visibile per posizionare la barra al centro verticale
        const visibleTop = Math.max(rect.top, 0);
        const visibleBottom = Math.min(rect.bottom, window.innerHeight);
        const visibleHeight = visibleBottom - visibleTop;
        const barHeight = ghost.offsetHeight || 12;
        const top = visibleTop + (visibleHeight / 2) - (barHeight / 2);
        ghost.style.top = top + 'px';
    }

    // Sincronizza la barra fantasma con lo scroll del contenitore
    ghost.addEventListener('scroll', () => {
        container.scrollLeft = ghost.scrollLeft;
    });
    // Sincronizza lo scroll del contenitore con la barra fantasma
    container.addEventListener('scroll', () => {
        ghost.scrollLeft = container.scrollLeft;
        updateInnerWidth();
    });
    // Registra eventi di scroll e ridimensionamento della finestra per riposizionare la barra
    document.addEventListener('scroll', repositionGhost);
    window.addEventListener('resize', repositionGhost);
    // Imposta larghezza iniziale e posizionamento
    updateInnerWidth();
    setTimeout(repositionGhost, 200);
}


/**
 * VERSIONE CORRETTA - Funzione mancante aggiunta
 * Raccoglie i dati da una singola riga della tabella dei dispositivi medici.
 */
function getMedicalDeviceRowData(row) {
    const cells = row.cells;
    return {
        // Struttura aggiornata: [Data, Codice, Descrizione, Cliente, Lotto, Quantità (con suffisso PZ)]
        data: cells[0].querySelector('input').value,
        codice: cells[1].querySelector('input').value,
        descrizione: cells[2].querySelector('input').value,
        cliente: cells[3].querySelector('input').value,
        lotto: cells[4].querySelector('input').value,
        // Quantità è nella colonna 5; numero scatoloni (ip.) è nella colonna 6
        quantita: cells[5] ? cells[5].querySelector('input').value : '',
        scatoloniTeorici: cells[6] ? cells[6].querySelector('input').value : ''
    };
}

/**
 * Restituisce i dati di produzione medical device attualmente presenti.
 * Se i dati sono memorizzati nel localStorage (medicalProductionData), li restituisce; altrimenti
 * raccoglie i dati dalla tabella.  Questa funzione viene utilizzata durante il salvataggio dei
 * dati sul server.
 */
function getMedicalProductionData() {
    // Prima prova a leggere dal localStorage
    try {
        const stored = localStorage.getItem('medicalProductionData');
        if (stored) {
            return JSON.parse(stored);
        }
    } catch (e) {
        console.warn('Impossibile leggere medicalProductionData dal localStorage:', e);
    }
    // In mancanza di dati persistiti, raccoglie quelli attualmente mostrati in tabella
    const data = [];
    const body = document.querySelector('#medicalDeviceProductionTable tbody');
    if (body) {
        body.querySelectorAll('tr').forEach(row => {
            const cells = row.querySelectorAll('td');
            // Struttura attuale: Data, Codice, Descrizione, Cliente, Lotto, Quantità (es. "5000 PZ")
            if (cells.length >= 6) {
                const qtyCellVal = cells[5].querySelector('input').value || '';
                // Parsa il numero di pezzi (senza suffisso), normalizzando i separatori decimali
                const numericQty = parseFloat(String(qtyCellVal).replace(/\./g, '').replace(',', '.')) || 0;
                data.push({
                    data: cells[0].querySelector('input').value,
                    codice: cells[1].querySelector('input').value,
                    descrizione: cells[2].querySelector('input').value,
                    cliente: cells[3].querySelector('input').value,
                    lotto: cells[4].querySelector('input').value,
                    quantita: numericQty,
                    unita: 0
                });
            }
        });
    }
    return data;
}

/**
 * VERSIONE CORRETTA - Funzione mancante aggiunta
 * Raccoglie e restituisce tutti i dati presenti nella tabella "Produzione Medical Device".
 */
function getAllMedicalDeviceData() {
    const data = [];
    const tableBody = document.querySelector('#medicalDeviceProductionTable tbody');
    if (tableBody) {
        tableBody.querySelectorAll('tr').forEach(row => {
            data.push(getMedicalDeviceRowData(row));
        });
    }
    return data;
}

    function updateDailyProductionTable() {
        dailyProductionTableBody.innerHTML = '';

        const allMainTableData = Array.from(productionTableBody.querySelectorAll('tr'))
            .map(row => getRowData(row));

        const filterCol = filterDailyColumnSelect.value;
        const filterVal = filterDailyValueInput.value.trim().toLowerCase();
        dailyProductionOperatorFilter = (filterCol === 'operatore' && filterVal) ? filterVal : 'Tutti';
        const tasksForDate = [];
        allMainTableData.forEach(rowData => {
            const prodDate = parseDateValue(rowData.produzioneData);
            const packDate = parseDateValue(rowData.dataConfezionamento);
            const dailyFilterDateStr = dailyProductionSelectedDate ? dailyProductionSelectedDate.toLocaleDateString('it-IT') : null;

            const isProductionTask = dailyFilterDateStr && prodDate === dailyFilterDateStr && parseFloat(rowData.quantitaDaProdurre) > 0;
            const isPackagingTask = dailyFilterDateStr && packDate === dailyFilterDateStr;

            if (isProductionTask && (!filterCol || !filterVal || String(rowData.codice || '').toLowerCase().includes(filterVal) || String(rowData.prodotto || '').toLowerCase().includes(filterVal) || String(rowData.cliente || '').toLowerCase().includes(filterVal) || String(rowData.operatore || '').toLowerCase().includes(filterVal))) {
                tasksForDate.push({
                    type: 'production',
                    data: rowData
                });
            }
            if (isPackagingTask && (!filterCol || !filterVal || String(rowData.codiceConfezionamento || '').toLowerCase().includes(filterVal) || String(rowData.prodotto || '').toLowerCase().includes(filterVal) || String(rowData.cliente || '').toLowerCase().includes(filterVal) || String(rowData.operatore || '').toLowerCase().includes(filterVal))) {
                tasksForDate.push({
                    type: 'packaging',
                    data: rowData
                });
            }
        });

        // Ordina i compiti per macchina (produzione vs confezionamento) e poi inseriscili.
        const sortedTasks = tasksForDate.sort((a, b) => {
            const machineA = getDailyProductionMachine(a.data, a.type);
            const machineB = getDailyProductionMachine(b.data, b.type);
            if (machineA === 'Confezionamento' && machineB !== 'Confezionamento') return 1;
            if (machineA !== 'Confezionamento' && machineB === 'Confezionamento') return -1;
            return (machineA || '').localeCompare(machineB || '');
        });
        sortedTasks.forEach(task => {
            dailyProductionTableBody.appendChild(createDailyProductionRow(task.data, task.type));
        });
        // Dopo aver popolato la tabella, aggiorna automaticamente le colonne OPE/OV e l'operatore
        if (typeof updateDailyOpeOv === 'function') {
            try {
                updateDailyOpeOv();
            } catch (err) {
                console.warn('Errore durante l\'aggiornamento automatico delle colonne OPE/OV:', err);
            }
        }
        // Aggiorna i suggerimenti di operatori e la lista dei macchinari dopo aver creato tutte le righe
        if (typeof updateOperatorSuggestions === 'function') {
            try {
                updateOperatorSuggestions();
            } catch (err) {
                console.warn('Errore durante l\'aggiornamento dei suggerimenti operatori:', err);
            }
        }
        if (typeof updateMacchinariOptionsListDaily === 'function') {
            try {
                updateMacchinariOptionsListDaily();
            } catch (err) {
                console.warn('Errore durante l\'aggiornamento della lista macchinari giornaliera:', err);
            }
        }
        // Rendi nuovamente ordinabile la tabella del programma giornaliero dopo l'aggiornamento
        if (typeof makeTableSortable === 'function') {
            makeTableSortable(document.getElementById('dailyProductionTable'));
        }
    }

    addDailyRowBtn.addEventListener('click', async () => {
        const newRow = createDailyProductionRow({});
        dailyProductionTableBody.appendChild(newRow);
        // Aggiorna i suggerimenti per operatori e macchinari quando si aggiunge una nuova riga
        if (typeof updateOperatorSuggestions === 'function') {
            try { updateOperatorSuggestions(); } catch (e) { console.warn('Errore aggiornamento suggerimenti operatori:', e); }
        }
        if (typeof updateMacchinariOptionsListDaily === 'function') {
            try { updateMacchinariOptionsListDaily(); } catch (e) { console.warn('Errore aggiornamento lista macchinari:', e); }
        }
        addLogEntry(`Aggiunta riga vuota al programma giornaliero.`);
        autoSaveAllData();
    });

    duplicateDailyRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('.daily-row-selector:checked');
        if (selectedRows.length === 0) {
            await showAlert('Seleziona almeno una riga da duplicare nel programma giornaliero.');
            return;
        }
        const confirmed = await showConfirm(`Sei sicuro di voler duplicare ${selectedRows.length} riga/e selezionata/e nel programma giornaliero?`);
        if (confirmed) {
            selectedRows.forEach(checkbox => {
                const originalRow = checkbox.closest('tr');
                const rowData = getDailyRowData(originalRow);
                const rowType = originalRow.classList.contains('packaging-row-bg') || originalRow.classList.contains('packaging-4xxxx-bg') ? 'packaging' : 'production';
                const newRow = createDailyProductionRow(rowData, rowType);
                originalRow.after(newRow);
            });
            // Aggiorna i suggerimenti dopo la duplicazione
            if (typeof updateOperatorSuggestions === 'function') {
                try { updateOperatorSuggestions(); } catch (e) { console.warn('Errore aggiornamento suggerimenti operatori:', e); }
            }
            if (typeof updateMacchinariOptionsListDaily === 'function') {
                try { updateMacchinariOptionsListDaily(); } catch (e) { console.warn('Errore aggiornamento lista macchinari:', e); }
            }
            await showAlert('Riga/e duplicata/e con successo nel programma giornaliero.');
            addLogEntry(`Duplicate ${selectedRows.length} riga/e nel programma giornaliero.`);
            autoSaveAllData();
        }
    });

    deleteDailyRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('.daily-row-selector:checked');
        if (selectedRows.length === 0) {
            await showAlert('Seleziona almeno una riga da eliminare nel programma giornaliero.');
            return;
        }
        const confirmed = await showConfirm(`Sei sicuro di voler eliminare ${selectedRows.length} riga/e selezionata/e dal programma giornaliero?`);
        if (confirmed) {
            selectedRows.forEach(checkbox => {
                checkbox.closest('tr').remove();
            });
            await showAlert('Riga/e eliminata/e con successo dal programma giornaliero.');
            addLogEntry(`Eliminate ${selectedRows.length} riga/e dal programma giornaliero.`);
            autoSaveAllData();
        }
    });

    // VERSIONE AGGIORNATA
// Attiva i filtri in tempo reale
// Quando si cambia la colonna di filtro, aggiorna il datalist per gli operatori
filterDailyColumnSelect.addEventListener('change', () => {
    const col = filterDailyColumnSelect.value;
    if (col === 'operatore') {
        // Usa il datalist per suggerire i nomi degli operatori
        filterDailyValueInput.setAttribute('list', 'operatorSuggestionsList');
        if (typeof updateOperatorSuggestions === 'function') {
            try { updateOperatorSuggestions(); } catch (e) { console.warn('Errore aggiornamento suggerimenti operatori:', e); }
        }
    } else {
        // Rimuovi il datalist se la colonna non è Operatore
        filterDailyValueInput.removeAttribute('list');
    }
    applyDailyFilter();
});
filterDailyValueInput.addEventListener('input', applyDailyFilter);
clearDailyFilterBtn.addEventListener('click', () => {
    filterDailyColumnSelect.value = '';
    filterDailyValueInput.value = '';
    dailyProductionOperatorFilter = 'Tutti';
    updateDailyProductionTable(); // Usa update per ricaricare la tabella senza filtri
});
    clearDailyFilterBtn.addEventListener('click', () => {
        filterDailyColumnSelect.value = '';
        filterDailyValueInput.value = '';
        dailyProductionOperatorFilter = 'Tutti';
        updateDailyProductionTable();
    });

    function applyDailyFilter() {
        const filterCol = filterDailyColumnSelect.value;
        const filterVal = filterDailyValueInput.value.trim().toLowerCase();
        dailyProductionOperatorFilter = (filterCol === 'operatore' && filterVal) ? filterVal : 'Tutti';

        dailyProductionTableBody.querySelectorAll('tr').forEach(row => {
            const rowData = getDailyRowData(row);
            let showRow = true;
            if (filterCol && filterVal && !String(rowData[filterCol] || '').toLowerCase().includes(filterVal)) {
                showRow = false;
            }
            row.style.display = showRow ? '' : 'none';
        });
    }

    saveDailyDataBtn.addEventListener('click', async () => {
        const saveName = await showPromptModal('Salva Programma Giornaliero', 'Inserisci un nome per questo salvataggio:', `programma_giornaliero_${new Date().toLocaleDateString('it-IT').replace(/\//g, '.')}`);
        if (saveName) {
            try {
                const allDailyData = Array.from(dailyProductionTableBody.querySelectorAll('tr')).map(row => getDailyRowData(row));
                const dailyDataKey = `daily_production_data_${saveName}`;
                localStorage.setItem(dailyDataKey, JSON.stringify(allDailyData));
                addLogEntry(`Programma giornaliero salvato manualmente: "${saveName}".`);
                await showAlert(`Programma giornaliero salvato con successo come "${saveName}"!`);
            } catch (e) {
                console.error("Errore durante il salvataggio del programma giornaliero:", e);
                addLogEntry(`Errore salvataggio manuale programma giornaliero: ${e.message}.`);
                await showAlert(`Errore durante il salvataggio del programma giornaliero: ${e.message}.`);
            }
        } else if (saveName === null) {
            await showAlert('Operazione di salvataggio annullata.');
        }
    });

    loadDailyDataBtn.addEventListener('click', async () => {
        const savedKeys = Object.keys(localStorage).filter(key => key.startsWith('daily_production_data_'));
        const savedNames = savedKeys.map(key => key.replace('daily_production_data_', ''));

        if (savedNames.length === 0) {
            await showAlert('Nessun programma giornaliero salvato trovato.');
            return;
        }

        const selectedAction = await showSelectionModal('Carica Programma Giornaliero', 'Seleziona un salvataggio:', savedNames);

        if (selectedAction === null) {
            await showAlert('Operazione di caricamento annullata.');
            return;
        }

        if (selectedAction === 'delete') {
            const selectedSaveName = document.getElementById('modalSelect').value;
            const confirmedDelete = await showConfirm(`Sei sicuro di voler eliminare il programma giornaliero "${selectedSaveName}"?`);
            if (confirmedDelete) {
                localStorage.removeItem(`daily_production_data_${selectedSaveName}`);
                addLogEntry(`Programma giornaliero "${selectedSaveName}" eliminato.`);
                await showAlert(`Programma giornaliero "${selectedSaveName}" eliminato.`);
                const remainingKeys = Object.keys(localStorage).filter(key => key.startsWith('daily_production_data_'));
                if (remainingKeys.length > 0) {
                    loadDailyDataBtn.click();
                } else {
                    dailyProductionTableBody.innerHTML = '';
                }
            }
            return;
        }

        const selectedSaveName = selectedAction;
        try {
            const key = `daily_production_data_${selectedSaveName}`;
            const savedData = localStorage.getItem(key);
            if (savedData) {
                const parsedData = JSON.parse(savedData);
                dailyProductionTableBody.innerHTML = '';
                parsedData.forEach(rowData => {
                    dailyProductionTableBody.appendChild(createDailyProductionRow(rowData));
                });
                await showAlert(`Programma giornaliero caricato con successo da "${selectedSaveName}".`);
                addLogEntry(`Programma giornaliero caricato: "${selectedSaveName}".`);
            } else {
                await showAlert('Programma giornaliero non trovato.');
            }
        } catch (e) {
            console.error("Errore durante il caricamento del programma giornaliero:", e);
            addLogEntry(`Errore caricamento programma giornaliero: ${e.message}.`);
            await showAlert(`Errore durante il caricamento del programma giornaliero: ${e.message}.`);
        }
    });

    
exportDailyPdfBtn.addEventListener('click', () => {
    // Abbreviazione dei nomi degli operatori per la stampa: cognome + iniziali.
    // Salviamo i valori originali in modo da ripristinarli dopo la stampa.
    const _operatorInputsForPrint = Array.from(document.querySelectorAll('#dailyProductionTable .col-daily-operatori input'));
    const _originalOperatorNames = _operatorInputsForPrint.map(inp => inp.value);
    _operatorInputsForPrint.forEach((inp, idx) => {
        inp.value = abbreviateOperatorName(inp.value);
    });
    // 1. Prende la data attualmente selezionata nel filtro
    const programDateForPrint = dailyProductionSelectedDate ?
        dailyProductionSelectedDate.toLocaleDateString('it-IT', {
            weekday: 'long',
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        }) :
        'Data non selezionata'; // Messaggio di default se nessuna data è scelta

    // 2. Inserisce la data nel punto giusto dell'intestazione nascosta
    document.getElementById('print-program-date').textContent = programDateForPrint;
    // 2bis. Inserisce la descrizione del filtro (se applicato) nell'intestazione di stampa
    try {
        const filterCol = (document.getElementById('filterDailyColumn') || {}).value || '';
        const filterVal = (document.getElementById('filterDailyValue') || {}).value || '';
        let filterText = '';
        if (filterCol && filterVal) {
            const colLabelMap = {
                codice: 'Codice',
                prodotto: 'Prodotto',
                cliente: 'Cliente',
                operatore: 'Operatore'
            };
            const friendlyCol = colLabelMap[filterCol] || filterCol;
            filterText = `Filtro: ${friendlyCol} = ${filterVal}`;
        }
        const printFilterElem = document.getElementById('print-filter-info');
        if (printFilterElem) {
            printFilterElem.textContent = filterText;
        }
    } catch (fErr) {
        console.warn('Errore nel calcolo del filtro per stampa:', fErr);
    }

    // 3. Aggiunge la classe per attivare gli stili di stampa corretti
    document.body.classList.add('printing-daily-production');

    // 4. Pulisce tutto dopo che la stampa è finita
    window.onafterprint = () => {
        document.body.classList.remove('printing-daily-production');
        window.onafterprint = null;
        // Ripristina i nomi degli operatori originali dopo la stampa
        _operatorInputsForPrint.forEach((inp, idx) => {
            inp.value = _originalOperatorNames[idx];
        });
    };

    // 5. Avvia la stampa
    setTimeout(() => {
        window.print();
    }, 100);
});

    exportDailyWordBtn.addEventListener('click', async () => {
        const data = Array.from(dailyProductionTableBody.querySelectorAll('tr')).map(row => getDailyRowData(row));
        if (data.length === 0) {
            await showAlert('Nessun dato nel programma giornaliero da esportare.');
            return;
        }

        const headers = [
            "Codice", "Prodotto", "Cliente", "Quantità", "Macchinario", "Quantità Confezionamento",
            "Operazioni", "Operatore", "Esito", "Quantità Prodotta", "Lotto", "TU", "TS", "Data Avallo"
        ];

        let htmlContent = `
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body { font-family: sans-serif; }
                    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
                    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; font-size: 0.9em; }
                    th { background-color: #f2f2f2; }
                    .production-row-bg { background-color: #F8F8FF; }
                    .packaging-row-bg { background-color: #F0FFF0; }
                    /* NUOVO STILE PER CASELLA ESITO */
                    .esito-box {
                        display: inline-block;
                        width: 20px;
                        height: 20px;
                        border: 1px solid #ccc;
                        cursor: pointer;
                        vertical-align: middle;
                    }
                    .esito-box.checked {
                        background-color: black;
                    }

* NUOVI STILI PER LARGHEZZA COLONNE E STAMPA SPEDIZIONI (con percentuali) */
#shippingScheduleTable {
    table-layout: fixed; /* Fondamentale per rispettare le larghezze */
    width: 100%;
    min-width: 2200px; /* Imposta una larghezza minima totale per lo scroll orizzontale */
}
#shippingScheduleTable th, #shippingScheduleTable td {
    white-space: normal;
    word-wrap: break-word;
}
#shippingScheduleTable th:nth-child(2) { width: 2%; }  /* OV */
#shippingScheduleTable th:nth-child(3) { width: 6%; }  /* Codice Articolo */
#shippingScheduleTable th:nth-child(4) { width: 40%; } /* Descrizione (più spazio) */
#shippingScheduleTable th:nth-child(5) { width: 4%; }  /* Quantità */
#shippingScheduleTable th:nth-child(6) { width: 3%; }  /* UM */
#shippingScheduleTable th:nth-child(7) { width: 6%; }  /* Data Consegna */
#shippingScheduleTable th:nth-child(8) { width: 6%; }  /* Data Conferma */
#shippingScheduleTable th:nth-child(9) { width: 12%; } /* Ragione Sociale */
#shippingScheduleTable th:nth-child(10){ width: 12%; } /* Rif. Cliente */
#shippingScheduleTable th:nth-child(11){ width: 15%; } /* Indirizzo */
#shippingScheduleTable th:nth-child(12){ width: 4%; }  /* CAP */
#shippingScheduleTable th:nth-child(13){ width: 6%; }  /* Città */
#shippingScheduleTable th:nth-child(14){ width: 2%; }  /* Prov */
#shippingScheduleTable th:nth-child(15){ width: 6%; }  /* Telefono */
@media print {
    body:not(.printing-shipping) #shippingScheduleContainer {
        display: none !important;
    }

    body.printing-shipping .sticky-controls-wrapper,
    body.printing-shipping header,
    body.printing-shipping .table-container,
    body.printing-shipping .sales-order-container,
    body.printing-shipping .gantt-chart-container,
    body.printing-shipping .daily-production-container:not(#shippingScheduleContainer),
    body.printing-shipping #logbookContainer,
    body.printing-shipping #analisiContainer {
        display: none !important;
    }

    body.printing-shipping #shippingScheduleContainer {
        display: block !important;
        box-shadow: none; border: none; padding: 0; margin: 0;
    }

    body.printing-shipping #shippingScheduleContainer .daily-production-controls {
        display: none;
    }
}

/* Stile per il raggruppamento degli OV nel Gantt Magazzino */
.gantt-ov-group {
    background-color: rgba(13, 71, 161, 0.85); /* Blu scuro semi-trasparente */
    border: 1px solid #1976D2;
    border-radius: 6px;
    padding: 4px;
    margin-bottom: 3px;
    width: 100%;
    box-sizing: border-box;
}
.gantt-ov-group-header {
    color: white;
    font-weight: bold;
    font-size: 0.8em;
    text-align: center;
    margin-bottom: 2px;
    border-bottom: 1px solid #64B5F6;
    padding-bottom: 2px;
}
/* Colore giallo intenso per i medical device nel Gantt spedizioni */
.gantt-task.shipping-task.medical-device-shipping {
    background-color: #FFD600; /* Giallo intenso */
    color: #424242; /* Testo più scuro per contrasto */
    border: 1px solid #FFAB00;
}

/* === STILI TOOLTIP DIVISO PER SPEDIZIONI === */
.tooltip-container {
    display: flex;
    gap: 8px;
    padding: 5px;
}
.tooltip-box {
    padding: 10px 14px;
    border-radius: 8px;
    border: 2px solid;
    min-width: 280px;
    font-size: 0.9em;
    background-color: #fff;
}
.tooltip-box h3 {
    margin: 0 0 8px 0;
    padding-bottom: 5px;
    border-bottom: 1px solid;
    font-size: 1.1em;
}
.tooltip-box p { margin: 0; line-height: 1.6; }
.tooltip-box strong { font-weight: 700; }

.shipping-info-tooltip {
    background-color: #E3F2FD; /* Blu chiaro - stile produzione */
    border-color: #64B5F6;
    color: #0D47A1;
}
.shipping-info-tooltip h3 { border-bottom-color: #90CAF9; color: #1565C0; }
.shipping-info-tooltip strong { color: #1976D2; }

.shipping-contact-tooltip {
    background-color: #E8F5E9; /* Verde chiaro - stile produzione */
    border-color: #81C784;
    color: #1B5E20;
}
.shipping-contact-tooltip h3 { border-bottom-color: #A5D6A7; color: #2E7D32; }
.shipping-contact-tooltip strong { color: #388E3C; }


/* =================================================================== */
/* ==> BLOCCO CSS PER GANTT MAGAZZINO E LAYOUT (CORRETTO) <== */
/* =================================================================== */

/* 1. Allarga il Gantt e il suo contenitore per occupare più spazio */
#warehouseGanttChartContainer {
    max-width: none;
    width: 100%;
    /* Abilita lo scorrimento orizzontale sul contenitore del Gantt di magazzino (definizione duplicata)
       per garantire che la proprietà overflow-x rimanga attiva anche in questa sezione. */
    /* overflow-x:auto removed by QBAR */
    overflow-y: hidden;
}
.warehouse-gantt-chart {
    /* Aggiornato per riflettere la larghezza totale delle 30 colonne (30 × 110px) più l'intestazione.
       Questo impedisce che il grafico venga schiacciato ma consente comunque di visualizzare tutte le
       colonne in un'unica schermata. */
    min-width: 3500px; /* 200px intestazione + 30×110px = 3500px */
}

/* 2. Stile per i gruppi di SPEDIZIONE (blu) */
.gantt-ov-group.shipping-group {
    background-color: #E3F2FD; /* Sfondo azzurro chiaro */
    border: 2px solid #0D47A1; /* Bordo blu scuro */
}
.gantt-ov-group.shipping-group .gantt-ov-group-header {
    background-color: #0D47A1; /* Intestazione blu scuro */
    color: white;
    border-bottom: 1px solid #64B5F6;
}

/* 3. Stile per i gruppi di ARRIVO (verde) */
.gantt-ov-group.arrival-group {
    background-color: #E8F5E9; /* Sfondo verde chiaro */
    border: 2px solid #1B5E20; /* Bordo verde scuro */
}
.gantt-ov-group.arrival-group .gantt-ov-group-header {
    background-color: #1B5E20; /* Intestazione verde scuro */
    color: white;
    border-bottom: 1px solid #81C784;
}

/* 4. Stile per le icone di layout nel Gantt */
.layout-icon {
    font-size: 0.9em;
    font-weight: bold;
    margin-left: 8px;
    color: #FFEB3B; /* Giallo per alta visibilità */
    display: inline-block;
}
.layout-icon .thermo-red {
    color: #E53935; /* Rosso per il termometro */
}
.layout-icon .snow-blue {
    color: #81D4FA; /* Azzurro per la neve */
}


/* =================================================================== */
/* ==> BLOCCO CSS UNIFICATO E DEFINITIVO PER TABELLE SPEDIZIONI E ARRIVI <== */
/* =================================================================== */

/* --- 1. Stili FONDAMENTALI Comuni per Tutte le Tabelle --- */
#shippingScheduleTable,
#arrivalScheduleTable,
#overdueArrivalsTable    
width: 100%;
    border-collapse: collapse;
    table-layout: fixed !important;
}

#shippingScheduleTable th, #shippingScheduleTable td,
#arrivalScheduleTable th, #arrivalScheduleTable td,
**#overdueArrivalsTable th, #overdueArrivalsTable td** { /* <-- AGGIUNTO QUI */
    padding: 4px 3px;
    text-align: center;
    border-bottom: 1px solid #eee;
    white-space: normal !important;
    word-wrap: break-word !important;
    vertical-align: middle;
    box-sizing: border-box;
}

#shippingScheduleTable th,
#arrivalScheduleTable th,
**#overdueArrivalsTable th** { /* <-- AGGIUNTO QUI */
    background-color: #CFD8DC;
    font-weight: bold;
    color: #333;
    position: sticky;
    top: 0;
    z-index: 5;
}

/* Stile per gli input, identico a quello della tabella di produzione principale */
#shippingScheduleTable input, #shippingScheduleTable select,
#arrivalScheduleTable input, #arrivalScheduleTable select,
**#overdueArrivalsTable input, #overdueArrivalsTable select** { /* <-- AGGIUNTO QUI */
    padding: 6px;
    border: 1px solid #cce7f0;
    border-radius: 5px;
    box-sizing: border-box;
    font-size: 0.85em;
    width: 100%;
    text-align: center;
}

/* Allineamento a sinistra solo per le colonne di testo lunghe */
#shippingScheduleTable td:nth-child(4) input,
#shippingScheduleTable td:nth-child(9) input,
#shippingScheduleTable td:nth-child(10) input,
#shippingScheduleTable td:nth-child(11) input,
#arrivalScheduleTable td:nth-child(4) input,
#arrivalScheduleTable td:nth-child(10) input,
#arrivalScheduleTable td:nth-child(11) input,
#arrivalScheduleTable td:nth-child(12) input,
**#overdueArrivalsTable td:nth-child(4) input,** /* <-- AGGIUNTE QUESTE 4 RIGHE */
**#overdueArrivalsTable td:nth-child(10) input,**
**#overdueArrivalsTable td:nth-child(11) input,**
**#overdueArrivalsTable td:nth-child(12) input** {
    text-align: left;
}

/* --- 2. Larghezze Specifiche e Identiche per le Colonne --- */
#shippingScheduleTable { min-width: 2300px; }
#arrivalScheduleTable, #overdueArrivalsTable { min-width: 2450px; } /* <-- AGGIUNTO QUI */

/* Colonne Comuni */
#shippingScheduleTable th:nth-child(1), #arrivalScheduleTable th:nth-child(1), **#overdueArrivalsTable th:nth-child(1)** { width: 30px; }
#shippingScheduleTable th:nth-child(2), #arrivalScheduleTable th:nth-child(2), **#overdueArrivalsTable th:nth-child(2)** { width: 80px; }
#shippingScheduleTable th:nth-child(3), #arrivalScheduleTable th:nth-child(3), **#overdueArrivalsTable th:nth-child(3)** { width: 100px; }
#shippingScheduleTable th:nth-child(4), #arrivalScheduleTable th:nth-child(4), **#overdueArrivalsTable th:nth-child(4)** { width: 430px; }

/* Colonna Layout (solo per tabelle Arrivi e Non Arrivati) */
#arrivalScheduleTable th:nth-child(5), **#overdueArrivalsTable th:nth-child(5)** { width: 150px; }

/* Colonne restanti (con indici sfalsati) */
#shippingScheduleTable th:nth-child(5),  #arrivalScheduleTable th:nth-child(6),  **#overdueArrivalsTable th:nth-child(6)** { width: 90px; }
#shippingScheduleTable th:nth-child(6),  #arrivalScheduleTable th:nth-child(7),  **#overdueArrivalsTable th:nth-child(7)** { width: 50px; }
#shippingScheduleTable th:nth-child(7),  #arrivalScheduleTable th:nth-child(8),  **#overdueArrivalsTable th:nth-child(8)** { width: 90px; }
#shippingScheduleTable th:nth-child(8),  #arrivalScheduleTable th:nth-child(9),  **#overdueArrivalsTable th:nth-child(9)** { width: 90px; }
#shippingScheduleTable th:nth-child(9),  #arrivalScheduleTable th:nth-child(10), **#overdueArrivalsTable th:nth-child(10)** { width: 300px; }
#shippingScheduleTable th:nth-child(10), #arrivalScheduleTable th:nth-child(11), **#overdueArrivalsTable th:nth-child(11)** { width: 200px; }
#shippingScheduleTable th:nth-child(11), #arrivalScheduleTable th:nth-child(12), **#overdueArrivalsTable th:nth-child(12)** { width: 320px; }
#shippingScheduleTable th:nth-child(12), #arrivalScheduleTable th:nth-child(13), **#overdueArrivalsTable th:nth-child(13)** { width: 70px; }
#shippingScheduleTable th:nth-child(13), #arrivalScheduleTable th:nth-child(14), **#overdueArrivalsTable th:nth-child(14)** { width: 140px; }
#shippingScheduleTable th:nth-child(14), #arrivalScheduleTable th:nth-child(15), **#overdueArrivalsTable th:nth-child(15)** { width: 50px; }
#shippingScheduleTable th:nth-child(15), #arrivalScheduleTable th:nth-child(16), **#overdueArrivalsTable th:nth-child(16)** { width: 140px; }


/* --- 3. Stili Specifici per la Tabella Merce Non Arrivata --- */
/* Questi stili rimangono separati perché sono unici per questa tabella. */
#overdueArrivalsTable tbody tr {
    background-color: #FFEBEE !important; /* Sfondo rossastro molto chiaro */
}
#overdueArrivalsTable tbody tr:hover {
    background-color: #FFCDD2 !important; /* Leggermente più scuro al passaggio del mouse */
}
#overdueArrivalsTable input {
    font-weight: normal;
    color: black;
}


/* Colore giallo più scuro per i medical device specifici in spedizione */
.gantt-task.shipping-task.medical-device-shipping-priority {
    background-color: #FFC107; /* Giallo più scuro/ambra */
    color: #000000; /* Testo nero per leggibilità */
    font-weight: bold;
    border: 2px solid #FF8F00;
}

/* Stile per l'icona di priorità (puntino rosso lampeggiante) */
.priority-icon {
    width: 10px;
    height: 10px;
    background-color: red;
    border-radius: 50%;
    display: inline-block;
    margin-right: 5px;
    animation: blink-animation 1s infinite;
    vertical-align: middle;
}

@keyframes blink-animation {
    0% { opacity: 1; }
    50% { opacity: 0.2; }
    100% { opacity: 1; }
}



/* Stile per l'icona di priorità (puntino rosso lampeggiante) */
.priority-icon {
    width: 10px;
    height: 10px;
    background-color: red;
    border-radius: 50%;
    display: inline-block;
    margin-right: 5px; /* Spazio tra il punto e il testo "OV:" */
    vertical-align: middle; /* Allinea verticalmente il punto al testo */
    animation: blink-animation 1s infinite;
}

@keyframes blink-animation {
    0% { opacity: 1; }
    50% { opacity: 0.2; }
    100% { opacity: 1; }
}

/* Colore giallo più scuro per i medical device specifici in spedizione */
.gantt-task.shipping-task.medical-device-shipping-priority {
    background-color: #FFC107; /* Giallo più scuro/ambra */
    color: #000000; /* Testo nero per leggibilità */
    font-weight: bold;
    border: 2px solid #FF8F00;
}

/* Stile per l'icona di priorità (puntino rosso lampeggiante) */
.priority-icon {
    width: 10px;
    height: 10px;
    background-color: red;
    border-radius: 50%;
    display: inline-block;
    margin-right: 5px;
    animation: blink-animation 1s infinite;
    vertical-align: middle;
}

@keyframes blink-animation {
    0% { opacity: 1; }
    50% { opacity: 0.2; }
    100% { opacity: 1; }
}



#shippingScheduleTable th, #shippingScheduleTable td,
#arrivalScheduleTable th, #arrivalScheduleTable td {
    padding: 4px 3px;
    text-align: center;
    border-bottom: 1px solid #eee;
    white-space: normal !important;   /* FONDAMENTALE per mandare il testo a capo */
    word-wrap: break-word !important; /* FONDAMENTALE per parole lunghe */
    vertical-align: middle;
    box-sizing: border-box;
}

#shippingScheduleTable th,
#arrivalScheduleTable th {
    background-color: #CFD8DC; /* Colore uniforme per le intestazioni */
    font-weight: bold;
    color: #333;
    position: sticky;
    top: 0;
    z-index: 5;
}



/* Colore giallo più scuro per i medical device specifici in spedizione */
.gantt-task.shipping-task.medical-device-shipping-priority {
    background-color: #FFC107; /* Giallo più scuro/ambra */
    color: #000000; /* Testo nero per leggibilità */
    font-weight: bold;
    border: 2px solid #FF8F00;
}

/* Stile per l'icona di priorità (puntino rosso lampeggiante) */
.priority-icon {
    width: 10px;
    height: 10px;
    background-color: red;
    border-radius: 50%;
    display: inline-block;
    margin-right: 5px;
    animation: blink-animation 1s infinite;
    vertical-align: middle;
}

@keyframes blink-animation {
    0% { opacity: 1; }
    50% { opacity: 0.2; }
    100% { opacity: 1; }
}


/* === STILI PER ICONE GANTT ARRIVI === */
.gantt-arrival-icon {
    display: inline-block;
    margin-right: 6px;
    font-weight: bold;
    vertical-align: middle;
}
.md-icon {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    color: black;
    background-color: #E0E0E0;
    padding: 1px 4px;
    border-radius: 3px;
    font-size: 0.8em;
    border: 1px solid #BDBDBD;
}

/* === STILI PER STELLE NEL GANTT (AGGIORNATI E PIÙ VISIBILI) === */
.gantt-star-icon {
    font-size: 1.7em; /* Aumentata la dimensione per maggiore visibilità */
    font-weight: bold; /* Aggiunto grassetto per farle risaltare */
    margin-right: 6px;
    vertical-align: middle;
    line-height: 1;
    /* Aggiunta una leggera ombra per migliorare il contrasto */
    text-shadow: 1px 1px 2px rgba(0,0,0,0.2); 
}
.yellow-star {
    color: #FFD600; /* Giallo più brillante e saturo */
}
.blue-star {
    color: #2196F3; /* Blu più vibrante e chiaro */
}

/* Mostra l'intestazione personalizzata durante la stampa del programma giornaliero */
body.printing-daily-production #print-header-info {
    display: block !important;
    text-align: center;
    border-bottom: 2px solid #333;
    padding-bottom: 10px;
    margin-bottom: 15px;
}

/* Nasconde i controlli e i filtri non necessari */
body.printing-daily-production .daily-production-controls {
    display: none !important;
}

/* =================================================================== */
/* ==> NUOVI STILI PER COMMENTI QA NELLE SPEDIZIONI <== */
/* =================================================================== */

/* Stile per il contenitore principale del tooltip che ora può avere 3 box */
.tooltip-container {
    display: flex;
    gap: 8px;
    padding: 5px;
    align-items: stretch; /* Assicura che i box abbiano la stessa altezza */
}

/* Stile per il nuovo box dei commenti QA */
.qa-comments-tooltip {
    background-color: #FFFDE7; /* Giallo molto chiaro */
    border-color: #FFD54F;
    color: #4E342E;
    display: flex;
    flex-direction: column; /* Organizza contenuto verticalmente */
}
.qa-comments-tooltip h3 {
    border-bottom-color: #FFF176;
    color: #BF360C;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.qa-comments-tooltip p {
    flex-grow: 1; /* Fa in modo che il paragrafo occupi lo spazio disponibile */
    white-space: pre-wrap; /* Mantiene la formattazione del testo (a capo, etc.) */
    text-align: left;
}

.lock-icon {
    cursor: pointer;
    font-size: 1.2em;
    margin-left: 10px;
    /* --- Righe aggiunte per allargare l'area --- */
    padding: 5px; /* Aggiunge spazio interno cliccabile */
    border-radius: 5px; /* Arrotonda leggermente gli angoli */
    transition: background-color 0.2s ease; /* Transizione fluida */
}

.lock-icon:hover {
    background-color: rgba(0, 0, 0, 0.1); /* Leggero sfondo al passaggio del mouse */
}

/* Stile per il modale di inserimento password e commenti */
.qa-modal-content {
    background-color: #fff;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
    text-align: center;
    max-width: 450px;
    width: 90%;
}
.qa-modal-content h3 {
    color: #BF360C;
    margin-top: 0;
}
.qa-modal-content p {
    color: #555;
    margin-bottom: 15px;
}
.qa-modal-content input,
.qa-modal-content textarea {
    width: calc(100% - 24px);
    padding: 12px;
    margin-bottom: 15px;
    border: 1px solid #ddd;
    border-radius: 8px;
    font-size: 1em;
    font-family: 'Quicksand', sans-serif;
}
.qa-modal-content textarea {
    min-height: 120px;
    resize: vertical;
    text-align: left;
}
.qa-modal-buttons {
    display: flex;
    justify-content: center;
    gap: 15px;
}

                </style>

            </head>
            <body>
                <h1>Programma Giornaliero di Produzione</h1>
                <p>Data Programma: ${dailyProductionSelectedDate ? dailyProductionSelectedDate.toLocaleDateString('it-IT') : 'N/D'}</p>
                <p>Operatore Filtrato: ${dailyProductionOperatorFilter || 'Tutti'}</p>
                <p>Data di Generazione: ${new Date().toLocaleDateString('it-IT')} ${new Date().toLocaleTimeString('it-IT')}</p>
                <table>
                    <thead>
                        <tr>
                            ${headers.map(h => `<th>${h}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${data.map(row => `
                            <tr class="${row.type === 'packaging' ? 'packaging-row-bg' : 'production-row-bg'}">
                                <td>${row.codice || ''}</td>
                                <td>${row.prodotto || ''}</td>
                                <td>${row.cliente || ''}</td>
                                <td>${row.quantita || ''}</td>
                                <td>${row.macchinario || ''}</td>
                                <td>${row.operazioni || ''}</td>
                                <td>${row.operatore || ''}</td>
                                <td>${row.esito || ''}</td>
                                <td>${row.quantitaProdotta || ''}</td>
                                <td>${row.lotto || ''}</td>
                                <td>${row.tu || ''}</td>
                                <td>${row.ts || ''}</td>
                                <td>${row.dataAvallo || ''}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </body>
            </html>
        `;

        const blob = new Blob([htmlContent], { type: 'application/msword;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        const todayForFileName = new Date().toLocaleDateString('it-IT').replace(/\//g, '.');

        let filename = `programma_giornaliero_${todayForFileName}`;
        if (dailyProductionSelectedDate) {
            filename += `_${dailyProductionSelectedDate.toLocaleDateString('it-IT').replace(/\//g, '.')}`;
        }
        if (dailyProductionOperatorFilter && dailyProductionOperatorFilter !== 'Tutti') {
            filename += `_operatore-${dailyProductionOperatorFilter.replace(/[^a-zA-Z0-9]/g, '')}`;
        }
        filename += `.doc`;

        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        addLogEntry(`Programma giornaliero esportato come Word: "${link.download}".`);
    });

    function createSalesOrderRow(rowData = {}) {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><input type="checkbox" class="sales-order-row-selector"></td>

            <td class="col-ov-flag"><div class="ov-flag-icon">&nbsp;</div></td>

            <td class="col-ov"><input type="text" value="${rowData.ov || ''}"></td>
            <td class="col-ov-codice"><input type="text" value="${rowData.codice || ''}"></td>
            <td class="col-ov-descrizione"><input type="text" value="${rowData.descrizione || ''}"></td>
            <td class="col-ov-quantita"><input type="number" value="${rowData.quantitaOrdine || ''}" min="0"></td>
            <td class="col-ov-um"><input type="text" value="${rowData.unitaMisura || ''}"></td>
            <td class="col-ov-data-consegna"><input type="text" class="datepicker" value="${rowData.dataConsegna || ''}"></td>
            <td class="col-ov-data-richiesta-cliente"><input type="text" class="datepicker" value="${rowData.dataRichiestaCliente || ''}"></td>
            <td class="col-ov-data-conferma"><input type="text" class="datepicker" value="${rowData.dataConferma || ''}"></td>
            <td class="col-ov-note"><input type="text" class="notes-input" value="${rowData.note || ''}"></td>
        `;

        row.querySelectorAll('.datepicker').forEach(input => flatpickr(input, {
            dateFormat: "d/m/Y",
            locale: "it"
        }));
        return row;
    }


    function getSalesOrderRowData(row) {
        return {
            ov: row.querySelector('.col-ov input').value,
            codice: row.querySelector('.col-ov-codice input').value,
            descrizione: row.querySelector('.col-ov-descrizione input').value,
            quantitaOrdine: row.querySelector('.col-ov-quantita input').value,
            unitaMisura: row.querySelector('.col-ov-um input').value,
            dataConsegna: row.querySelector('.col-ov-data-consegna input').value,
            dataRichiestaCliente: row.querySelector('.col-ov-data-richiesta-cliente input').value,
            dataConferma: row.querySelector('.col-ov-data-conferma input').value,
            note: row.querySelector('.col-ov-note input').value
        };
    }

// ========================================================================
    // ==> FUNZIONI NUOVE PER LA TABELLA MEDICAL DEVICE <==
    // ========================================================================

    /**
     * Crea una riga HTML per la tabella dei dispositivi medici.
     */
    function createMedicalDeviceRow(rowData = {}, isManual = false) {
        const row = document.createElement('tr');
        
        let siringhePerScatola = '';
        const confezionamentoString = rowData.rawConfezionamentoDettaglio || 
                                      (rowData.confezionamentoPezzi ? `${rowData.confezionamentoPezzi}x` : '');

        const match = confezionamentoString.match(/^(\d)x/);
        if (match && (match[1] === '1' || match[1] === '2')) {
            siringhePerScatola = match[1];
        }
        
        const volumeProduzione = rowData.quantitaDaProdurre ? `${rowData.quantitaDaProdurre} Kg` : '';

        const isEditable = (currentUserLevel === 3 || currentUserLevel === 6);
        const readOnlyAttr = isEditable ? '' : 'readonly';

        row.innerHTML = `
            <td><input type="text" value="${rowData.codice || ''}" readonly></td>
            <td><input type="text" value="${rowData.prodotto || ''}" readonly></td>
            <td><input type="text" value="${rowData.cliente || ''}" readonly></td>
            <td><input type="text" value="${rowData.confezionamentoPezzi || ''}" readonly></td>
            <td><input type="text" class="scarti-input" value="${rowData.scarti || ''}" ${readOnlyAttr}></td>
            <td><input type="text" value="${volumeProduzione}" readonly></td>
            <td><input type="text" value="${siringhePerScatola}" readonly></td>
        `;
        
        if(isManual){
             row.querySelectorAll('input').forEach(input => {
                if(!input.classList.contains('scarti-input')) {
                    input.readOnly = false;
                }
             });
        }
        
        return row;
    }

    /**
     * Aggiorna e filtra la tabella dei dispositivi medici basandosi sulla tabella principale.
     */
    function updateMedicalDeviceProductionTable() {
        if (!medicalDeviceTableBody) return;
        
        const scartiValues = new Map();
        medicalDeviceTableBody.querySelectorAll('tr').forEach(row => {
            const codice = row.cells[0].querySelector('input').value;
            const scarti = row.cells[4].querySelector('input').value;
            if (codice && scarti) {
                scartiValues.set(codice, scarti);
            }
        });

        medicalDeviceTableBody.innerHTML = '';
        const allMainTableData = getAllTableData();
        
        const medicalDevicesData = allMainTableData.filter(row => isMedicalDeviceCode(row.codice));

        const startDate = medicalDeviceStartDateInput._flatpickr.selectedDates[0];
        const endDate = medicalDeviceEndDateInput._flatpickr.selectedDates[0];
        if (endDate) endDate.setHours(23, 59, 59, 999);

        const filterCodiceText = filterMedicalDeviceCodice.value.toLowerCase();
        const filterDescrizioneText = filterMedicalDeviceDescrizione.value.toLowerCase();
        const filterClienteText = filterMedicalDeviceCliente.value.toLowerCase();

        const filteredMedicalData = medicalDevicesData.filter(row => {
            const prodDateParts = row.produzioneData.split('/');
            if (prodDateParts.length !== 3) return false;
            const rowDate = new Date(parseInt(prodDateParts[2]), parseInt(prodDateParts[1]) - 1, parseInt(prodDateParts[0]));
            const dateMatch = (!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate);
            if (!dateMatch) return false;

            const codiceMatch = !filterCodiceText || (row.codice || '').toLowerCase().includes(filterCodiceText);
            const descrizioneMatch = !filterDescrizioneText || (row.prodotto || '').toLowerCase().includes(filterDescrizioneText);
            const clienteMatch = !filterClienteText || (row.cliente || '').toLowerCase().includes(filterClienteText);

            return codiceMatch && descrizioneMatch && clienteMatch;
        });

        filteredMedicalData.forEach(rowData => {
            if (scartiValues.has(rowData.codice)) {
                rowData.scarti = scartiValues.get(rowData.codice);
            }
            medicalDeviceTableBody.appendChild(createMedicalDeviceRow(rowData));
        });
    }

    addSalesOrderRowBtn.addEventListener('click', () => {
        salesOrderTableBody.appendChild(createSalesOrderRow());
        addLogEntry(`Aggiunta riga manuale all'Ordine di Vendita.`);
        autoSaveAllData();
    });

    duplicateSalesOrderRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('.sales-order-row-selector:checked');
        if (selectedRows.length === 0) {
            await showAlert('Seleziona almeno una riga OV da duplicare.');
            return;
        }
        const confirmed = await showConfirm(`Sei sicuro di voler duplicare ${selectedRows.length} riga/e selezionata/e?`);
        if (confirmed) {
            selectedRows.forEach(checkbox => {
                const originalRow = checkbox.closest('tr');
                const rowData = getSalesOrderRowData(originalRow);
                const newRow = createSalesOrderRow(rowData);
                originalRow.after(newRow);
            });
            await showAlert('Riga/e duplicata/e con successo.');
            addLogEntry(`Duplicate ${selectedRows.length} riga/e OV.`);
            autoSaveAllData();
        }
    });
    deleteSalesOrderRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('.sales-order-row-selector:checked');
        if (selectedRows.length === 0) {
            await showAlert('Seleziona almeno una riga OV da eliminare.');
            return;
        }
        const confirmed = await showConfirm(`Sei sicuro di voler eliminare ${selectedRows.length} riga/e OV selezionata/e?`);
        if (confirmed) {
            selectedRows.forEach(checkbox => {
                const rowToDelete = checkbox.closest('tr');
                const rowData = getSalesOrderRowData(rowToDelete);
                addLogEntry(`Eliminata riga OV (OV: ${rowData.ov || 'N/D'}, Codice: ${rowData.codice || 'N/D'}).`);
                rowToDelete.remove();
            });
            await showAlert('Riga/e eliminata/e con successo.');
            autoSaveAllData();
        }
    });
    sendEmailOVBtn.addEventListener('click', () => {
        const recipient = 'rossella.crippa@iralab.it';
        const subject = encodeURIComponent('Riepilogo e Aggiornamento Ordini di Vendita');
        const tableRowsData = getAllSalesOrderData();
        const currentDateTime = new Date().toLocaleString('it-IT');

        let body = `Gentile Rossella,\n\n`;
        body += `in allegato un riepilogo degli ordini di vendita aggiornato al ${currentDateTime}.\n\n`;
        body += `--- RIEPILOGO ORDINI ---\n`;

        if (tableRowsData.length > 0) {
            tableRowsData.forEach((row, index) => {
                body += `\n${index + 1}. OV: ${row.ov || 'N/D'}\n`;
                body += `    Codice: ${row.codice || 'N/D'}\n`;
                body += `    Descrizione: ${row.descrizione || 'N/D'}\n`;
                body += `    Quantità: ${row.quantitaOrdine || 'N/D'} ${row.unitaMisura || ''}\n`;
                body += `    Data Richiesta Cliente: ${row.dataRichiestaCliente || 'N/D'}\n`;
                body += `    Data Spedizione Pianificata: ${row.dataConsegna || 'N/D'}\n`;
            });
        } else {
            body += `\nNessun ordine di vendita presente al momento.\n`;
        }

        body += `\n\nCordiali saluti,\nSistema di Programmazione Produzione`;

        const mailtoLink = `mailto:${recipient}?subject=${subject}&body=${encodeURIComponent(body)}`;

        try {
            window.location.href = mailtoLink;
            addLogEntry(`Email Ordini di Vendita inviata a ${recipient}.`);
        } catch (e) {
            addLogEntry(`Errore invio email OV: ${e.message}.`);
            showAlert(`Impossibile aprire il client di posta. Assicurati che un client di posta sia configurato.`);
        }
    });

    async function handlePossibleDuplicateOV(rowData) {
        const existingRows = Array.from(salesOrderTableBody.querySelectorAll('tr'));

        const isDuplicate = existingRows.some(row => {
            const existingRowData = getSalesOrderRowData(row);
            // Controlla se i campi chiave sono identici
            return (existingRowData.ov.trim() === rowData.ov.trim() &&
                existingRowData.codice.trim() === rowData.codice.trim() &&
                existingRowData.descrizione.trim() === rowData.descrizione.trim());
        });

        if (isDuplicate) {
            // Se trova un duplicato, chiede conferma all'utente
            return await showConfirm(
                'Una riga con lo stesso OV, Codice e Descrizione esiste già. Vuoi inserirla comunque?',
                'Duplicato Rilevato'
            );
        }

        // Se non è un duplicato, ritorna true per procedere con l'inserimento
        return true;
    }
    function setOVFlag(row, status) {
        const flagCell = row.querySelector('.col-ov-flag div');
        if (!flagCell) return;

        let icon = '&nbsp;';
        let tooltipText = '';

        switch (status) {
            // Regola 4
            case 'NOT_FOUND':
            case 'MISMATCH':
                icon = `<div class="ov-flag-icon red-triangle">&#9650;<span class="exclamation-mark">!</span></div>`;
                tooltipText = 'Prodotto non presente o dati non corrispondenti nel programma di produzione.';
                break;
                // Regola 3
            case 'NO_MATERIALS':
                icon = `<div class="ov-flag-icon dark-yellow-square">&#9632;<span class="exclamation-mark">!</span></div>`;
                tooltipText = 'Articolo da produrre senza materie prime.';
                break;
                // Regole 1 e 2 (prima dell'eliminazione)
            case 'DELETE_OK_STOCK':
            case 'DELETE_OK_PROD':
                icon = `<div class="ov-flag-icon green-square">&#9632;</div>`;
                tooltipText = 'Articolo pronto per la spedizione. La riga verrà rimossa.';
                break;
                // Nessun flag
            case 'CLEAR':
            default:
                icon = '&nbsp;';
                tooltipText = '';
                break;
        }

        flagCell.innerHTML = icon;
        flagCell.parentElement.title = tooltipText;
    }

    // PRIMA FUNZIONE DA SOSTITUIRE
    function runFullCheck() {
    setInterval(autoSaveAllData, 120000);
        const allOVRows = Array.from(salesOrderTableBody.querySelectorAll('tr'));
        const rowsToDelete = [];

        allOVRows.forEach(row => {
            const status = checkSalesOrderRowStatus(row); // Ottiene lo stato dalla funzione di controllo
            setOVFlag(row, status); // Imposta il flag in base allo stato

            // Se lo stato indica che la riga va eliminata, la aggiunge alla lista
            if (status === 'DELETE_OK_STOCK' || status === 'DELETE_OK_PROD') {
                rowsToDelete.push(row);
            }
        });

        // Se ci sono righe da eliminare, le rimuove e mostra un solo messaggio riepilogativo
        if (rowsToDelete.length > 0) {
            rowsToDelete.forEach(row => row.remove());
            showAlert(`Sono state verificate ed eliminate automaticamente ${rowsToDelete.length} righe OV perché la loro richiesta è stata soddisfatta.`);
        }

        // Salva sempre lo stato attuale della tabella OV
        autoSaveAllData();
    }

    // SECONDA FUNZIONE DA SOSTITUIRE
    function checkSalesOrderRowStatus(salesOrderRow) {
        const ovData = getSalesOrderRowData(salesOrderRow);
        const ovCode = ovData.codice.trim();

        if (!ovCode) {
            return 'CLEAR'; // Nessun codice, nessun flag
        }

        const allProdRowsData = getAllTableData();
        const matchingProdRows = allProdRowsData.filter(pRow => pRow.codiceConfezionamento === ovCode);

        // Regola 4 (parziale): Codice non trovato nel programma di produzione
        if (matchingProdRows.length === 0) {
            return 'NOT_FOUND';
        }

        // Regola 5: Cerca la corrispondenza migliore basata sulla data
        const bestMatchProdRow = matchingProdRows.find(pRow =>
            pRow.dataSpedizione === ovData.dataRichiestaCliente
        ) || matchingProdRows[0]; // Se non trova la data, usa la prima corrispondenza

        const quantityMatch = parseFloat(bestMatchProdRow.confezionamentoPezzi) === parseFloat(ovData.quantitaOrdine);
        const dateMatch = bestMatchProdRow.dataSpedizione === ovData.dataRichiestaCliente;

        // Regola 4 (completa): Se codice, quantità o data non corrispondono
        if (!quantityMatch || !dateMatch) {
            return 'MISMATCH';
        }

        // Se tutto corrisponde, controlliamo lo stato della produzione e dei materiali
        const needsProduction = parseFloat(bestMatchProdRow.giacenzaMagazzino) < parseFloat(bestMatchProdRow.quantitaRichiesta);

        // Regola 1: Corrispondenza perfetta E prodotto già a magazzino
        if (!needsProduction) {
            return 'DELETE_OK_STOCK';
        }

        // Se serve produrre, controlliamo i materiali
        const materialsOK = bestMatchProdRow.materiePrime === 'si' && bestMatchProdRow.materialeConfezionamento === 'si';

        // Regola 2: Corrispondenza perfetta, serve produrre E i materiali sono disponibili
        if (needsProduction && materialsOK) {
            return 'DELETE_OK_PROD';
        }

        // Regola 3: Corrispondenza perfetta, serve produrre MA i materiali NON sono disponibili
        if (needsProduction && !materialsOK) {
            return 'NO_MATERIALS';
        }

        return 'CLEAR'; // Caso di default, nessun flag
    }

    function runFullCheck() {
        const allOVRows = Array.from(salesOrderTableBody.querySelectorAll('tr'));
        const rowsToDelete = [];

        allOVRows.forEach(row => {
            const status = checkSalesOrderRowStatus(row); // Ottiene lo stato dalla funzione di controllo
            setOVFlag(row, status); // Imposta il flag in base allo stato

            // Se lo stato indica che la riga va eliminata, la aggiunge alla lista
            if (status === 'DELETE_OK_STOCK' || status === 'DELETE_OK_PROD') {
                rowsToDelete.push(row);
            }
        });

        // Se ci sono righe da eliminare, le rimuove e mostra un solo messaggio riepilogativo
        if (rowsToDelete.length > 0) {
            rowsToDelete.forEach(row => row.remove());
            showAlert(`Sono state verificate ed eliminate automaticamente ${rowsToDelete.length} righe OV perché la loro richiesta è stata soddisfatta.`);
        }

        // Salva sempre lo stato attuale della tabella OV
        autoSaveAllData();
    }

    dailyProductionFlatpickr = flatpickr(dailyProductionDateInput, {
        dateFormat: "d/m/Y",
        locale: "it",
        allowInput: true,
        defaultDate: today,
        onChange: (selectedDates) => {
            dailyProductionSelectedDate = selectedDates[0] || null;
            updateDailyProductionTable();
        }
    });
    dailyProductionSelectedDate = dailyProductionFlatpickr.selectedDates[0] || today;
    updateDailyProductionTable();

    // Nuove logiche per la tabella Analisi

    let analysisHeaders = [];
    let methodHeaders = [];
    let analysisPlan = {};
    let referenzeMap = {};

    function parseCsv(csvString) {
        const lines = csvString.split('\n').map(line => line.trim()).filter(line => line.length > 0);
        return lines.map(line => {
            const row = [];
            let inQuote = false;
            let currentField = '';
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                if (char === '"') {
                    inQuote = !inQuote;
                } else if (char === ',' && !inQuote) {
                    row.push(currentField.trim());
                    currentField = '';
                } else {
                    currentField += char;
                }
            }
            row.push(currentField.trim());
            return row;
        });
    }

    function loadAnalisiExcelData(pianoAnaliticoCsv, referenzeCsv) {
        const mdPianoAnaliticoRows = pianoAnaliticoCsv;
        const headerRowIndex = 4;
        const methodsRowIndex = 5;
        const dataStartIndex = 6;
        const dataColOffset = 4;

        analysisHeaders = mdPianoAnaliticoRows[headerRowIndex].slice(dataColOffset);
        methodHeaders = mdPianoAnaliticoRows[methodsRowIndex].slice(dataColOffset);

        analysisPlan = {};
        for (let i = dataStartIndex; i < mdPianoAnaliticoRows.length; i++) {
            const row = mdPianoAnaliticoRows[i];
            if (row.length < 2) continue;
            const productCode = String(row[1]).trim();
            const productName = String(row[2]).trim();

            if (!productCode) continue;

            analysisPlan[productCode] = { name: productName, analyses: {} };
            for (let j = 0; j < analysisHeaders.length; j++) {
                const analysisName = analysisHeaders[j];
                const cellContent = String(row[j + dataColOffset] || '').trim();
                if (cellContent !== '' && cellContent.toUpperCase() !== 'N.A.' && cellContent.toUpperCase() !== 'NA') {
                    analysisPlan[productCode].analyses[analysisName] = true;
                }
            }
        }

        referenzeMap = {};
        for (let i = 2; i < referenzeCsv.length; i++) {
            const row = referenzeCsv[i];
            if (row.length < 5) continue;
            const prodottoCapostipiteRef = String(row[1]).trim();
            const semiconfezionatoRef = String(row[2]).trim();
            const referenzaEquivalente = String(row[3]).trim();
            const referenzaEquivalenteRef = String(row[4]).trim();

            if (!referenzeMap[referenzaEquivalenteRef]) {
                referenzeMap[referenzaEquivalenteRef] = {};
            }
            if (!referenzeMap[referenzaEquivalenteRef].codiceFinito) {
                referenzeMap[referenzaEquivalenteRef].codiceFinito = prodottoCapostipiteRef;
            }
            if (!referenzeMap[referenzaEquivalenteRef].codiceSemi) {
                referenzeMap[referenzaEquivalenteRef].codiceSemi = semiconfezionatoRef;
            }
            if (!referenzeMap[referenzaEquivalenteRef].codiceBulk) {
                referenzeMap[referenzaEquivalenteRef].codiceBulk = semiconfezionatoRef.replace(/\*SC$/, '');
            }
            if (!referenzeMap[referenzaEquivalenteRef].nomeProdotto) {
                referenzeMap[referenzaEquivalenteRef].nomeProdotto = referenzaEquivalente;
            }
        }
    }

    function loadAnalisiStaticData() {
        try {
            const savedReferenze = localStorage.getItem('referenzeData');
            const savedPianoAnalitico = localStorage.getItem('pianoAnaliticoData');

            if (savedReferenze && savedPianoAnalitico) {
                const referenzeParsed = JSON.parse(savedReferenze);
                const pianoAnaliticoParsed = JSON.parse(savedPianoAnalitico);

                loadAnalisiExcelData(pianoAnaliticoParsed, referenzeParsed);

                // Aggiorna le etichette dei nomi file ma non mostra i flag, poiché utilizziamo soltanto la data dell'ultimo import.
                referenzeFileStatusSpan.textContent = `File: ${localStorage.getItem('referenzeFileName')}`;
                referenzeFileStatusSpan.style.display = 'none';
                pianoAnaliticoFileStatusSpan.textContent = `File: ${localStorage.getItem('pianoAnaliticoFileName')}`;
                pianoAnaliticoFileStatusSpan.style.display = 'none';

                console.log("Dati di analisi caricati da localStorage.");
            } else {
                console.log("Nessun dato di analisi trovato in localStorage.");
            }
        } catch (e) {
            console.error("Errore nel caricamento dei file di analisi da localStorage:", e);
        }
    }

    function saveAnalisiStaticData(key, data, filename) {
        try {
            localStorage.setItem(key, JSON.stringify(data));
            localStorage.setItem(`${key.replace('Data', 'FileName')}`, filename);
            console.log(`Dati per ${key} salvati in localStorage.`);
        } catch (e) {
            console.error(`Errore nel salvataggio dei dati per ${key} in localStorage:`, e);
        }
    }

    
function syncAnalysisStatuses(groupId, analysisName, newStatus, type) {
    // Sincronizza lo stato di una specifica analisi per tutte le righe appartenenti allo stesso gruppo e dello stesso tipo.
    if (!groupId || !analysisName || !type) return;
    // Trova l'indice della colonna dell'analisi basandosi sul nome
    const headerCells = Array.from(document.querySelectorAll('#analisiTable thead th .analysis-name'));
    const analysisIndex = headerCells.findIndex(el => el.textContent.trim() === analysisName);
    if (analysisIndex === -1) return;
    // Seleziona tutte le righe del gruppo con il tipo specificato
    const allRows = document.querySelectorAll(`#analisiTable tbody tr[data-group-id="${groupId}"][data-type="${type}"]`);
    allRows.forEach(row => {
        // Ogni cella di analisi è allineata con l'ordine delle intestazioni
        const analysisCells = row.querySelectorAll('td.analysis-cell');
        const cellToUpdate = analysisCells[analysisIndex];
        if (cellToUpdate) {
            const box = cellToUpdate.querySelector('.conformity-box');
            if (box) {
                box.classList.remove('green-flag', 'red-x');
                if (newStatus === 'conform') {
                    box.classList.add('green-flag');
                } else if (newStatus === 'non-conform') {
                    box.classList.add('red-x');
                }
            }
        }
        // Aggiorna lo stato della riga e delle righe madri del gruppo
        updateAnalysisRowStatus(row);
        updateParentRowStatus(row);
    });
}


// ===================================================================
    // ==> SOSTITUISCI LA VECCHIA FUNZIONE createConformityBox CON QUESTA <==
    // ===================================================================
   function createConformityBox(code, analysisName, lottoSC, type, groupId = null, initialStatus = 'unknown') {
    // Costruisce la chiave di stato usando il gruppo, oppure il lotto se non fornito.  
    // Includere anche il tipo per distinguere tra bulk e semi; questo consente di sincronizzare i semafori tra righe con lo stesso gruppo.
    const groupKey = groupId || lottoSC || 'no-lotto';
    const statusKey = `${groupKey}_${type}_${analysisName}`;
    const savedStatus = localStorage.getItem(statusKey) || initialStatus;
    
    const box = document.createElement('div');
    box.classList.add('conformity-box');
    if (savedStatus === 'conform') box.classList.add('green-flag');
    if (savedStatus === 'non-conform') box.classList.add('red-x');

    // QUESTA PARTE CONTROLLA SE L'UTENTE È CQ (5) O AMMINISTRATORE (6)
    if (currentUserLevel === 5 || currentUserLevel === 6) {
        
        // Se l'utente è autorizzato, rende il flag cliccabile
        box.style.cursor = 'pointer'; 
        
        box.addEventListener('click', (event) => {
            event.stopPropagation();
            let newStatus = 'unknown';

            if (box.classList.contains('green-flag')) {
                box.classList.remove('green-flag');
                box.classList.add('red-x');
                newStatus = 'non-conform';
            } else if (box.classList.contains('red-x')) {
                box.classList.remove('red-x');
                newStatus = 'unknown';
            } else {
                box.classList.add('green-flag');
                newStatus = 'conform';
            }
            
            localStorage.setItem(statusKey, newStatus);
            // Sincronizza i semafori fra tutte le righe con lo stesso gruppo e tipo
            syncAnalysisStatuses(groupKey, analysisName, newStatus, type);

            // Aggiorna lo stato visivo della riga
            const currentRow = box.closest('tr');
            if(currentRow) {
                updateAnalysisRowStatus(currentRow);
                updateParentRowStatus(currentRow);
            }
        });
    } else {
        // Per tutti gli altri utenti, il flag non è cliccabile
        box.style.cursor = 'not-allowed';
    }

    return box;
}

    function renderAnalysisTableHeaders() {
        const analysisTable = document.getElementById('analisiTable');
        const thead = analysisTable.querySelector('thead');
        thead.innerHTML = '';
        if (analysisHeaders.length === 0) return;

        const mainHeaderRow = document.createElement('tr');
        mainHeaderRow.innerHTML = `
    <th rowspan="2"></th>
    <th rowspan="2">Prodotto</th>
    <th rowspan="2">Lotto SC</th>
    <th rowspan="2">Data di Produzione</th>
    <th rowspan="2">Stato</th>
    <th colspan="${analysisHeaders.length}">Analisi Chimico-Fisiche e Microbiologiche</th>
`;
        thead.appendChild(mainHeaderRow);

        const analysisNamesRow = document.createElement('tr');
        let analysisCellsHtml = '';
        analysisHeaders.forEach((header, index) => {
            const method = methodHeaders[index] || '&nbsp;';
            analysisCellsHtml += `<th><div class="analysis-name">${header}</div><div class="method-column">${method}</div></th>`;
        });
        analysisNamesRow.innerHTML = analysisCellsHtml;
        thead.appendChild(analysisNamesRow);
    }

    


   
function toggleAnalysisRow(parentRow) {
    const icon = parentRow.querySelector('.rotate-icon');
    if (!icon) return; 
    
    const parentLotto = parentRow.dataset.parentLotto;
    const childrenRows = document.querySelectorAll(`.analisi-child-row[data-parent-lotto="${parentLotto}"]`);
    
    // Controlla il simbolo corretto ('-')
    const isExpanded = icon.textContent.trim() === '-';
    
    childrenRows.forEach(child => {
        child.style.display = isExpanded ? 'none' : 'table-row';
    });
    
    // Alterna tra '+' e '-'
    icon.textContent = isExpanded ? '+' : '-';
} 


   function updateAnalysisRowStatus(row) {
    const statusCell = row.querySelector('.status-cell .status-indicator, .status-cell .status-nc');
    if (!statusCell) return;

    const conformityBoxes = row.querySelectorAll('.conformity-box');
    
    if (conformityBoxes.length === 0) {
        statusCell.className = 'status-indicator';
        return;
    }

    let greenCount = 0;
    let redCount = 0;
    let totalBoxes = 0;
    
    // Controlla solo i semafori che hanno una classe di stato
    conformityBoxes.forEach(box => {
        totalBoxes++;
        if (box.classList.contains('green-flag')) {
            greenCount++;
        } else if (box.classList.contains('red-x')) {
            redCount++;
        }
    });

    statusCell.className = 'status-indicator'; // Resetta lo stato

    if (redCount > 0) {
        statusCell.className = 'status-nc';
        statusCell.textContent = '❌';
    } else if (greenCount === totalBoxes && totalBoxes > 0) {
        statusCell.classList.add('status-green');
        statusCell.textContent = '';
    } else if (greenCount > 0 || redCount > 0) {
        statusCell.classList.add('status-yellow');
        statusCell.textContent = '';
    } else {
        statusCell.classList.add('status-red');
        statusCell.textContent = '';
    }

    // Aggiorna lo stato della riga madre se la riga attuale è una figlia
    if (row.classList.contains('analisi-child-row')) {
        updateParentRowStatus(row);
    }
}
   
 function updateParentRowStatus(childRow) {
    // Determina il gruppo a cui appartiene la riga figlia. Utilizziamo il dataset.groupId per gestire
    // gruppi composti dallo stesso bulk e semi-confezionato, indipendentemente dal numero di lotto.
    const groupId = childRow.dataset.groupId;
    if (!groupId) return;
    // Trova tutte le righe madri (prodotto finito) appartenenti a questo gruppo
    const parentRows = document.querySelectorAll(`.collapsible-row[data-group-id="${groupId}"]`);
    if (!parentRows || parentRows.length === 0) return;
    // Trova tutte le righe figlie del gruppo (bulk e semi)
    const childrenRows = document.querySelectorAll(`.analisi-child-row[data-group-id="${groupId}"]`);
    let allChildrenGreen = true;
    let hasNonConform = false;
    let hasPending = false;
    let totalChildrenWithAnalyses = 0;
    childrenRows.forEach(child => {
        const childIndicator = child.querySelector('.status-indicator, .status-nc');
        if (childIndicator) {
            totalChildrenWithAnalyses++;
            if (childIndicator.classList.contains('status-nc')) {
                hasNonConform = true;
            } else if (childIndicator.classList.contains('status-green')) {
                // allChildrenGreen rimane true
            } else if (childIndicator.classList.contains('status-yellow')) {
                hasPending = true;
                allChildrenGreen = false;
            } else {
                allChildrenGreen = false;
            }
        }
    });
    // Aggiorna ciascuna riga madre del gruppo
    parentRows.forEach(parentRow => {
        const parentStatusIndicator = parentRow.querySelector('.status-indicator, .status-nc');
        if (!parentStatusIndicator) return;
        // Resetta lo stato
        parentStatusIndicator.className = 'status-indicator';
        parentStatusIndicator.textContent = '';
        if (totalChildrenWithAnalyses === 0) {
            parentStatusIndicator.classList.add('status-red');
        } else if (hasNonConform) {
            parentStatusIndicator.className = 'status-nc';
            parentStatusIndicator.textContent = '❌';
        } else if (allChildrenGreen) {
            parentStatusIndicator.classList.add('status-green');
        } else if (hasPending) {
            parentStatusIndicator.classList.add('status-yellow');
        } else {
            parentStatusIndicator.classList.add('status-red');
        }
    });
}

    
// VERSIONE AGGIORNATA
function createAnalisiRow(rowData) {
    const row = document.createElement('tr');
    // Utilizza il groupId passato se esiste; altrimenti usa lottoSC o una combinazione di codice e tipo.
    const computedGroupId = rowData.groupId || rowData.lottoSC || `${rowData.code}_${rowData.type}`;
    row.dataset.groupId = computedGroupId;
    row.dataset.uniqueId = rowData.uniqueId;
    row.dataset.code = rowData.code || '';
    row.dataset.type = rowData.type;
    row.dataset.parentCode = rowData.parentCode || rowData.code;
    row.dataset.lotto = rowData.lottoSC || '';
    row.dataset.parentLotto = rowData.parentLottoSC || rowData.lottoSC;

    if (rowData.isChild) {
        row.classList.add('analisi-child-row');
        row.style.display = 'none';
    } else if (!rowData.isUnplanned) {
        row.classList.add('collapsible-row');
    }

    let productCellHtml = '';
    let lottoCellValue = rowData.lottoSC || '';

    if (rowData.isChild) {
        productCellHtml = `<textarea class="analisi-input">${rowData.name} (${rowData.code})</textarea>`;
    } else if (rowData.isUnplanned) {
        productCellHtml = `<strong><textarea class="analisi-input">${rowData.name} (${rowData.code})</textarea></strong>`;
    } else {
        const finishedCodeWithAsterisk = isMedicalDeviceCode(rowData.code) ? `${rowData.code}*` : rowData.code;
        productCellHtml = `<span class="rotate-icon">+</span> <strong><textarea class="analisi-input">${rowData.name} (${finishedCodeWithAsterisk})</textarea></strong>`;
        lottoCellValue = '';
    }

    row.innerHTML = `
        <td>
            <input type="checkbox" class="analisi-row-selector">
            <span class="info-icon">(i)</span>
        </td>
        <td>${productCellHtml}</td>
        <td><input type="text" class="analisi-input" value="${lottoCellValue}"></td>
        <td><input type="text" class="analisi-input" value="${rowData.produzioneData || ''}"></td>
        <td class="status-cell"><span class="status-indicator"></span></td>
    `;
    
    const icon = row.querySelector('.rotate-icon');
    if (icon) {
        icon.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleAnalysisRow(row);
        });
    }

    row.querySelectorAll('input, textarea').forEach(input => {
        input.addEventListener('click', e => e.stopPropagation());
    });

    const infoIcon = row.querySelector('.info-icon');
    infoIcon.addEventListener('click', (e) => {
        e.stopPropagation();
        const allProdData = getAllTableData();
        const parentCode = row.dataset.parentCode || row.dataset.code;
        const parentLotto = row.dataset.parentLotto || row.dataset.lotto;
        let prodData = allProdData.find(pRow => pRow.codice === parentCode && pRow.lottoSC === parentLotto);
        if (prodData) {
            showSplitTooltip(prodData, e);
        } else {
            showAlert('Dati di produzione di riferimento non trovati.');
        }
    });

    const analysisChecks = analysisPlan[rowData.code] ? analysisPlan[rowData.code].analyses : {};
    const defaultUnplannedAnalyses = ['Stato Fisico', 'Colore', 'Odore'];

    // Per ogni intestazione di analisi crea una cella.  Utilizziamo il groupId della riga,
    // salvato in row.dataset.groupId, per sincronizzare i semafori tra le righe con lo stesso gruppo.
    const currentGroupId = row.dataset.groupId;
    analysisHeaders.forEach(header => {
        const cell = document.createElement('td');
        cell.classList.add('analysis-cell');
        // Mostra il quadratino se l'analisi è nel piano oppure, per righe manuali/non pianificate, se è una delle analisi di default.
        if (
            (analysisChecks[header]) ||
            (rowData.code === 'manual') ||
            (rowData.isUnplanned && defaultUnplannedAnalyses.includes(header))
        ) {
            cell.appendChild(createConformityBox(rowData.code, header, rowData.lottoSC, rowData.type, currentGroupId));
        }
        cell.addEventListener('contextmenu', async (e) => {
            e.preventDefault();
            const hasBox = cell.querySelector('.conformity-box');
            const actionText = hasBox ? 'rimuovere' : 'aggiungere';
            if (await showConfirm(`Vuoi ${actionText} l'analisi per questa cella?`)) {
                if (hasBox) {
                    hasBox.remove();
                } else {
                    cell.appendChild(createConformityBox(rowData.code, header, rowData.lottoSC, rowData.type, currentGroupId));
                }
                updateAnalysisRowStatus(row);
            }
        });
        row.appendChild(cell);
    });

    row.addEventListener('mouseout', hideGenericTooltip);

    row.querySelectorAll('textarea.analisi-input').forEach(textarea => {
        const autoResize = () => {
            textarea.style.height = 'auto';
            textarea.style.height = textarea.scrollHeight + 'px';
        };
        textarea.addEventListener('input', autoResize);
        setTimeout(autoResize, 0);
    });
    setTimeout(() => updateAnalysisRowStatus(row), 0);

    return row;
} 

    // ==> INCOLLA QUESTO BLOCCO DI CODICE QUI <==

    // NUOVA FUNZIONE PER AGGIORNARE GLI STATI AL CARICAMENTO
    function refreshAllAnalysisStatuses() {
        analisiTableBody.querySelectorAll('tr').forEach(row => {
            if (!row.classList.contains('collapsible-row')) { // Aggiorna prima i figli
                updateAnalysisRowStatus(row);
            }
        });
        analisiTableBody.querySelectorAll('.collapsible-row').forEach(row => { // E poi i genitori
            const firstChild = document.querySelector(`.analisi-child-row[data-parent-code="${row.dataset.code}"][data-parent-lotto="${row.dataset.lotto}"]`);
            if (firstChild) updateParentRowStatus(firstChild);
        });
    }

    // ===================================================================
    // ==> 2. SOSTITUISCI L'INTERA VECCHIA applyAnalisiFilters CON QUESTA <==
    // ===================================================================
    function applyAnalisiFilters() {
        if (!analisiSearchLottoInput || !analisiStartDateInput || !analisiEndDateInput) return;

        const lottoSearchText = analisiSearchLottoInput.value.trim().toLowerCase();
        const startDate = flatpickr.parseDate(analisiStartDateInput.value, "d/m/Y");
        const endDate = flatpickr.parseDate(analisiEndDateInput.value, "d/m/Y");
        if (endDate) {
            endDate.setHours(23, 59, 59, 999);
        }

        const allParentRows = analisiTableBody.querySelectorAll('tr.collapsible-row');

        allParentRows.forEach(parentRow => {
            const parentDateInput = parentRow.querySelector('td:nth-child(4) input');
            const rowDate = parentDateInput ? flatpickr.parseDate(parentDateInput.value, "d/m/Y") : null;

            // Condizione 1: La data della riga madre deve corrispondere
            let dateMatch = true;
            if (startDate && endDate) {
                dateMatch = rowDate ? (rowDate >= startDate && rowDate <= endDate) : false;
            }

            // Condizione 2: Il lotto cercato deve essere presente in ALMENO una delle righe del gruppo
            let lottoMatch = !lottoSearchText;
            if (lottoSearchText) {
                const gruppoId = parentRow.dataset.parentLotto.toLowerCase();
                const childrenRows = document.querySelectorAll(`.analisi-child-row[data-parent-lotto="${gruppoId}"]`);

                // Controlla il lotto della riga madre (che corrisponde a quello del semiconfezionato)
                if (parentRow.dataset.lotto.toLowerCase().includes(lottoSearchText)) {
                    lottoMatch = true;
                }

                // Se non ha ancora trovato, controlla le figlie
                if (!lottoMatch) {
                    for (const child of childrenRows) {
                        if (child.dataset.lotto.toLowerCase().includes(lottoSearchText)) {
                            lottoMatch = true;
                            break; // Trovato, inutile continuare a ciclare
                        }
                    }
                }
            }

            const showGroup = dateMatch && lottoMatch;

            // Applica la visibilità all'INTERO GRUPPO
            parentRow.style.display = showGroup ? '' : 'none';
            const childrenRows = document.querySelectorAll(`.analisi-child-row[data-parent-lotto="${parentRow.dataset.parentLotto}"]`);
            childrenRows.forEach(child => {
                child.style.display = showGroup ? '' : 'none';
            });
        });

        saveAnalisiFiltersState();
    }

    function clearAnalisiFilters() {
        document.getElementById('searchLottoInput').value = '';

        const startDatePicker = document.getElementById('analisiStartDate')._flatpickr;
        const endDatePicker = document.getElementById('analisiEndDate')._flatpickr;

        const today = new Date();
        const fourteenDaysLater = new Date();
        fourteenDaysLater.setDate(today.getDate() + 14);

        startDatePicker.setDate(today, false);
        endDatePicker.setDate(fourteenDaysLater, false);

        applyAnalisiFilters();
    }

    function refreshStaticDataFromStorage() {
        try {
            const savedReferenze = localStorage.getItem('referenzeData');
            const savedPianoAnalitico = localStorage.getItem('pianoAnaliticoData');

            if (savedReferenze && savedPianoAnalitico) {
                const referenzeParsed = JSON.parse(savedReferenze);
                const pianoAnaliticoParsed = JSON.parse(savedPianoAnalitico);

                loadAnalisiExcelData(pianoAnaliticoParsed, referenzeParsed);
                console.log("Dati statici (Referenze, Piano Analitico) ricaricati forzatamente dalla memoria.");
            }
        } catch (e) {
            console.error("Errore nel ricaricare i dati statici dalla memoria:", e);
        }
    }

    function updateAnalisiTable() {
    refreshStaticDataFromStorage();
    const tbody = analisiTableBody;
    if (!tbody) return;
    if (analysisHeaders.length === 0) {
        renderAnalysisTableHeaders();
    }
    tbody.innerHTML = '';
    const productionData = getAllTableData();

    // 1. Separa i prodotti pianificati da quelli non pianificati
    const plannedProducts = [];
    const unplannedProducts = [];

    productionData.forEach(prodRow => {
        if (prodRow.produzioneData && isMedicalDeviceCode(prodRow.codice)) {
            if (referenzeMap[prodRow.codice]) {
                plannedProducts.push(prodRow);
            } else {
                unplannedProducts.push(prodRow);
            }
        }
    });

    // 2. Renderizza i prodotti con un piano analitico (logica standard)
    plannedProducts.forEach((prodRow, index) => {
        const finitoCode = prodRow.codice;
        const referenza = referenzeMap[finitoCode];
        if (!referenza) { 
            console.error(`Referenza non trovata per il codice pianificato: ${finitoCode}`);
            return;
        }
        // Calcola una chiave di gruppo univoca basata su codice bulk e semi
        // In questo modo referenze con lo stesso bulk e semi (A e B) saranno considerate appartenenti allo stesso gruppo.
        const baseGroupKey = `${referenza.codiceBulk}_${referenza.codiceSemi}`;
        const gruppoId = `${baseGroupKey}_${prodRow.lottoSC || 'no-lotto'}`;
        // Riga principale (prodotto finito)
        tbody.appendChild(createAnalisiRow({
            code: finitoCode,
            name: prodRow.prodotto,
            lottoSC: prodRow.lottoSC,
            produzioneData: prodRow.produzioneData,
            type: 'finito',
            isChild: false,
            parentLottoSC: gruppoId,
            groupId: gruppoId
        }));
        // Riga semi-confezionato
        const semiName = analysisPlan[referenza.codiceSemi]?.name || `Semi-confezionato (${referenza.codiceSemi})`;
        tbody.appendChild(createAnalisiRow({
            code: referenza.codiceSemi,
            name: semiName,
            lottoSC: prodRow.lottoSC,
            produzioneData: prodRow.produzioneData,
            type: 'semi',
            isChild: true,
            parentCode: finitoCode,
            parentLottoSC: gruppoId,
            groupId: gruppoId
        }));
        // Riga bulk
        const bulkName = analysisPlan[referenza.codiceBulk]?.name || `Bulk (${referenza.codiceBulk})`;
        tbody.appendChild(createAnalisiRow({
            code: referenza.codiceBulk,
            name: bulkName,
            lottoSC: deriveBulkLotto(prodRow.lottoSC),
            produzioneData: prodRow.produzioneData,
            type: 'bulk',
            isChild: true,
            parentCode: finitoCode,
            parentLottoSC: gruppoId,
            groupId: gruppoId
        }));
    });

    // 3. Renderizza i prodotti senza piano analitico nella loro sezione speciale
    if (unplannedProducts.length > 0) {
        const colCount = 5 + analysisHeaders.length; // 5 colonne fisse + N colonne analisi
        const separatorRow = document.createElement('tr');
        separatorRow.innerHTML = `<th colspan="${colCount}" style="background-color: #ffcdd2; color: #b71c1c; text-align: center; font-size: 1.1em; padding: 10px; border-top: 2px solid #b71c1c; border-bottom: 2px solid #b71c1c;">Analisi senza Piano Analitico</th>`;
        tbody.appendChild(separatorRow);

        unplannedProducts.forEach(prodRow => {
            tbody.appendChild(createAnalisiRow({
                code: prodRow.codice,
                name: prodRow.prodotto,
                lottoSC: prodRow.lottoSC,
                produzioneData: prodRow.produzioneData,
                type: 'finito-unplanned',
                isChild: false,
                isUnplanned: true // Flag speciale per gestirle diversamente
            }));
        });
    }
    
    applyAnalisiFilters();
    setTimeout(refreshAllAnalysisStatuses, 150);
}


    // VERSIONE AGGIORNATA
addAnalisiRowBtn.addEventListener('click', async () => {
    const nome = await showPromptModal('Nuova Riga Analisi', 'Inserisci il nome del prodotto:', 'Prodotto...');
    if (!nome) return;
    const lotto = await showPromptModal('Nuova Riga Analisi', 'Inserisci il numero di lotto:', 'Lotto...');
    if (!lotto) return;

    // Crea una riga manuale contrassegnandola come "non pianificata"
    const newRow = createAnalisiRow({
        name: `${nome} (Manuale)`,
        lottoSC: lotto,
        code: 'manual', 
        isUnplanned: true // Aggiunge il flag per mostrare solo le 3 analisi di default
    });

    analisiTableBody.appendChild(newRow);
    addLogEntry(`Aggiunta riga di analisi manuale per ${nome} - ${lotto}.`);
});

// VERSIONE DEFINITIVA E SEMPLIFICATA
function createAnalisiRow(rowData) {
    const row = document.createElement('tr');
    // Calcola l'identificatore di gruppo. Se rowData.groupId è presente, utilizza quello;
    // in caso contrario usa il lotto o una combinazione di codice e tipo.
    const computedGroupId = rowData.groupId || rowData.lottoSC || `${rowData.code}_${rowData.type}`;
    row.dataset.groupId = computedGroupId;
    row.dataset.code = rowData.code || '';
    row.dataset.type = rowData.type;
    row.dataset.parentCode = rowData.parentCode || rowData.code;
    row.dataset.lotto = rowData.lottoSC || '';
    row.dataset.parentLotto = rowData.parentLottoSC || rowData.lottoSC;
    if (rowData.isUnplanned) {
        row.dataset.isUnplanned = 'true';
    }
    if (rowData.isChild) {
        row.classList.add('analisi-child-row');
        row.style.display = 'none';
    } else if (!rowData.isUnplanned) {
        row.classList.add('collapsible-row');
    }
    let productCellHtml = '';
    let lottoCellValue = rowData.lottoSC || '';
    if (rowData.isChild) {
        productCellHtml = `<textarea class="analisi-input">${rowData.name} (${rowData.code})</textarea>`;
    } else if (rowData.isUnplanned) {
        productCellHtml = `<strong><textarea class="analisi-input">${rowData.name}</textarea></strong>`;
    } else {
        const finishedCodeWithAsterisk = isMedicalDeviceCode(rowData.code) ? `${rowData.code}*` : rowData.code;
        productCellHtml = `<span class="rotate-icon">+</span> <strong><textarea class="analisi-input">${rowData.name} (${finishedCodeWithAsterisk})</textarea></strong>`;
        lottoCellValue = '';
    }
    row.innerHTML = `
        <td>
            <input type="checkbox" class="analisi-row-selector">
            <span class="info-icon">(i)</span>
        </td>
        <td>${productCellHtml}</td>
        <td><input type="text" class="analisi-input" value="${lottoCellValue}"></td>
        <td><input type="text" class="analisi-input" value="${rowData.produzioneData || ''}"></td>
        <td class="status-cell"><span class="status-indicator"></span></td>
    `;
    const icon = row.querySelector('.rotate-icon');
    if (icon) {
        icon.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleAnalysisRow(row);
        });
    }
    row.querySelectorAll('input, textarea').forEach(input => {
        input.addEventListener('click', e => e.stopPropagation());
    });
    const infoIcon = row.querySelector('.info-icon');
    infoIcon.addEventListener('click', (e) => {
        e.stopPropagation();
        const allProdData = getAllTableData();
        const parentCode = row.dataset.parentCode || row.dataset.code;
        const parentLotto = row.dataset.parentLotto || row.dataset.lotto;
        let prodData = allProdData.find(pRow => pRow.codice === parentCode && pRow.lottoSC === parentLotto);
        if (prodData) {
            showSplitTooltip(prodData, e);
        } else if (!rowData.isUnplanned) {
            // Non mostrare l'alert per le righe manuali/unplanned
            showAlert('Dati di produzione di riferimento non trovati.');
        }
    });
    const analysisChecks = analysisPlan[rowData.code] ? analysisPlan[rowData.code].analyses : {};
    const defaultUnplannedAnalyses = ['Stato Fisico', 'Colore', 'Odore'];
    analysisHeaders.forEach(header => {
        const cell = document.createElement('td');
        cell.classList.add('analysis-cell');
        if (
            analysisChecks[header] ||
            (rowData.isUnplanned && defaultUnplannedAnalyses.includes(header))
        ) {
            cell.appendChild(createConformityBox(rowData.code, header, rowData.lottoSC, rowData.type, computedGroupId));
        }
        cell.addEventListener('contextmenu', async (e) => {
            e.preventDefault();
            const hasBox = cell.querySelector('.conformity-box');
            const actionText = hasBox ? 'rimuovere' : 'aggiungere';
            if (await showConfirm(`Vuoi ${actionText} l'analisi per questa cella?`)) {
                if (hasBox) hasBox.remove();
                else cell.appendChild(createConformityBox(rowData.code, header, rowData.lottoSC, rowData.type, computedGroupId));
                updateAnalysisRowStatus(row);
            }
        });
        row.appendChild(cell);
    });
    row.addEventListener('mouseout', hideGenericTooltip);
    row.querySelectorAll('textarea.analisi-input').forEach(textarea => {
        const autoResize = () => {
            textarea.style.height = 'auto';
            textarea.style.height = textarea.scrollHeight + 'px';
        };
        textarea.addEventListener('input', autoResize);
        setTimeout(autoResize, 0);
    });
    setTimeout(() => updateAnalysisRowStatus(row), 0);
    return row;
}

    // VERSIONE CORRETTA
duplicateAnalisiRowBtn.addEventListener('click', async () => {
    const selectedRows = document.querySelectorAll('#analisiTable .analisi-row-selector:checked');
    if (selectedRows.length === 0) {
        await showAlert('Seleziona almeno una riga da duplicare nella tabella Analisi.');
        return;
    }
    const confirmed = await showConfirm(`Sei sicuro di voler duplicare ${selectedRows.length} riga/e selezionata/e nella tabella Analisi?`);
    if (confirmed) {
        selectedRows.forEach(checkbox => {
            const originalRow = checkbox.closest('tr');

            // Legge correttamente i dati dalla riga originale
            const productName = originalRow.querySelector('td:nth-child(2) textarea.analisi-input').value;
            const lottoSc = originalRow.querySelector('td:nth-child(3) input.analisi-input').value;
            const prodDate = originalRow.querySelector('td:nth-child(4) input.analisi-input').value;

            // Ricostruisce l'oggetto dati completo per la nuova riga
            const rowData = {
                name: productName,
                lottoSC: lottoSc,
                produzioneData: prodDate,
                code: originalRow.dataset.code,
                type: originalRow.dataset.type,
                isChild: originalRow.classList.contains('analisi-child-row'),
                isUnplanned: originalRow.hasAttribute('data-is-unplanned'),
                parentCode: originalRow.dataset.parentCode,
                parentLottoSC: originalRow.dataset.parentLotto
            };

            const newRow = createAnalisiRow(rowData);
            originalRow.after(newRow);
        });
        await showAlert('Riga/e duplicata/e con successo.');
        addLogEntry(`Duplicate ${selectedRows.length} riga/e nella tabella Analisi.`);
    }
});
// =================================================================
// ATTIVAZIONE EVENTI PER PROGRAMMA GIORNALIERO DI SPEDIZIONE
// =================================================================
if (importOSBtn) {
    importOSBtn.addEventListener('click', () => {
        // Utilizza un input file dinamico per gestire l'importazione degli OS in
        // maniera indipendente dalla logica generale di importMode.  In questo
        // modo si evita qualsiasi conflitto con altri gestori o il blocco dei
        // pointer events.
        const dynInput = document.createElement('input');
        dynInput.type = 'file';
        dynInput.accept = '.xls,.xlsx,.csv';
        dynInput.style.display = 'none';
        document.body.appendChild(dynInput);
        dynInput.addEventListener('change', async (evt) => {
            const file = evt.target.files[0];
            if (!file) {
                dynInput.remove();
                return;
            }
            const ext = file.name.split('.').pop().toLowerCase();
            const isExcel = ext === 'xls' || ext === 'xlsx';
            const isCsv = ext === 'csv';
            if (isExcel || isCsv) {
                await processOSFile(file, shippingScheduleTableBody, createShippingScheduleRow);
                // Rendi di nuovo ordinabile la tabella dopo l'importazione da file
                if (typeof makeTableSortable === 'function') {
                    makeTableSortable(document.getElementById('shippingScheduleTable'));
                }
                // Aggiorna timestamp dell'ultimo import OS
                const ts = formatDateTimeForDisplay(new Date());
                try {
                    // Usa una chiave uniforme senza underscore
                    localStorage.setItem('lastImportOS', ts);
                } catch (e) {}
                if (typeof updateImportTimestamps === 'function') updateImportTimestamps();
            } else {
                await showAlert('Formato file non supportato per OS. Seleziona un file Excel o CSV.');
            }
            dynInput.remove();
        });
        dynInput.click();
    });
}

if (addShippingRowBtn) {
    addShippingRowBtn.addEventListener('click', () => {
        // Aggiunge una nuova riga alla tabella spedizioni e aggiorna l'ordinamento
        shippingScheduleTableBody.appendChild(createShippingScheduleRow());
        // Rende nuovamente ordinabile la tabella dopo l'aggiunta di righe dinamiche
        if (typeof makeTableSortable === 'function') {
            makeTableSortable(document.getElementById('shippingScheduleTable'));
        }
        autoSaveAllData();
    });
}

if (duplicateShippingRowBtn) {
    duplicateShippingRowBtn.addEventListener('click', () => {
        const selectedRows = document.querySelectorAll('#shippingScheduleTable .shipping-row-selector:checked');
        selectedRows.forEach(checkbox => {
            const originalRow = checkbox.closest('tr');
            const rowData = getShippingScheduleRowData(originalRow);
            originalRow.after(createShippingScheduleRow(rowData));
        });
        // Aggiorna Gantt e rende di nuovo ordinabile la tabella spedizioni
        updateWarehouseGanttChart();
        if (typeof makeTableSortable === 'function') {
            makeTableSortable(document.getElementById('shippingScheduleTable'));
        }
        autoSaveAllData();
    });
}

if (deleteShippingRowBtn) {
    deleteShippingRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('#shippingScheduleTable .shipping-row-selector:checked');
        if (selectedRows.length > 0 && await showConfirm(`Eliminare ${selectedRows.length} righe di spedizione?`)) {
            selectedRows.forEach(checkbox => checkbox.closest('tr').remove());
            updateWarehouseGanttChart();
            // Aggiorna l'ordinabilità della tabella spedizioni dopo eliminazione
            if (typeof makeTableSortable === 'function') {
                makeTableSortable(document.getElementById('shippingScheduleTable'));
            }
            autoSaveAllData();
        }
    });
}

if (sendShippingEmailBtn) {
    sendShippingEmailBtn.addEventListener('click', () => {
        // Logica invio email (può essere implementata in dettaglio dopo)
        showAlert("Funzionalità 'Invia Mail' per le spedizioni da implementare.");
    });
}

if (exportShippingDataBtn) {
    exportShippingDataBtn.addEventListener('click', () => {
        // Logica per esportare in CSV la tabella spedizioni
        const data = getAllShippingData();
        if (data.length === 0) return;
        const headers = Object.keys(data[0]);
        let csvContent = headers.join(';') + '\n';
        data.forEach(row => {
            csvContent += headers.map(header => `"${row[header]}"`).join(';') + '\n';
        });
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `programma_spedizioni_${new Date().toLocaleDateString('it-IT').replace(/\//g, '-')}.csv`;
        link.click();
    });
}

// Filtri per la tabella spedizioni
const filterShippingColumn = document.getElementById('filterShippingColumn');
const filterShippingValue = document.getElementById('filterShippingValue');
const applyShippingFilterBtn = document.getElementById('applyShippingFilterBtn');
const clearShippingFilterBtn = document.getElementById('clearShippingFilterBtn');
const clearShippingDateBtn = document.getElementById('clearShippingDateBtn');


function applyShippingFilter() {
    if (!filterShippingColumn || !filterShippingValue || !shippingStartDateInput || !shippingEndDateInput) return;

    const filterCol = filterShippingColumn.value;
    const filterVal = filterShippingValue.value.trim().toLowerCase();
    const startDate = flatpickr.parseDate(shippingStartDateInput.value, "d/m/Y");
    const endDate = flatpickr.parseDate(shippingEndDateInput.value, "d/m/Y");
    if (endDate) endDate.setHours(23, 59, 59, 999); // Include tutto il giorno finale

    document.querySelectorAll('#shippingScheduleTable tbody tr').forEach(row => {
        let showRow = true;
        const rowData = getShippingScheduleRowData(row);

        // Filtro per testo
        if (filterCol && filterVal && !String(rowData[filterCol] || '').toLowerCase().includes(filterVal)) {
            showRow = false;
        }

        // Filtro per data di consegna
        if (showRow && startDate) {
            const rowDateParts = rowData.dataConsegna.split('/');
            if (rowDateParts.length === 3) {
                const rowDate = new Date(parseInt(rowDateParts[2]), parseInt(rowDateParts[1]) - 1, parseInt(rowDateParts[0]));
                if (rowDate < startDate || (endDate && rowDate > endDate)) {
                    showRow = false;
                }
            } else {
                showRow = false; // Nasconde righe senza data valida se un filtro data è attivo
            }
        }

        row.style.display = showRow ? '' : 'none';
    });
}


// --- VERSIONE AGGIORNATA PER SPEDIZIONI ---

// Attiva i filtri in tempo reale
if (filterShippingColumn) filterShippingColumn.addEventListener('change', applyShippingFilter);
if (filterShippingValue) filterShippingValue.addEventListener('input', applyShippingFilter);
if (clearShippingFilterBtn) {
    clearShippingFilterBtn.addEventListener('click', () => {
        if (filterShippingColumn) filterShippingColumn.value = '';
        if (filterShippingValue) filterShippingValue.value = '';
        applyShippingFilter();
    });
}

// Inizializzazione dei calendari con filtro "live"
if (shippingStartDateInput && shippingEndDateInput) {
    const oggi = new Date();
    const settimanaProssima = new Date();
    settimanaProssima.setDate(oggi.getDate() + 7);

    flatpickr(shippingStartDateInput, { 
        dateFormat: "d/m/Y", 
        locale: "it",
        defaultDate: oggi,
        onChange: applyShippingFilter // Filtra al cambio della data
    });
    flatpickr(shippingEndDateInput, { 
        dateFormat: "d/m/Y", 
        locale: "it",
        defaultDate: settimanaProssima,
        onChange: applyShippingFilter // Filtra al cambio della data
    });
    
    setTimeout(applyShippingFilter, 500);
}


if(applyShippingFilterBtn) {
    applyShippingFilterBtn.addEventListener('click', () => {
        const filterCol = filterShippingColumn.value;
        const filterVal = filterShippingValue.value.trim().toLowerCase();
        document.querySelectorAll('#shippingScheduleTable tbody tr').forEach(row => {
            const rowData = getShippingScheduleRowData(row);
            row.style.display = (filterCol && filterVal && !String(rowData[filterCol] || '').toLowerCase().includes(filterVal)) ? 'none' : '';
        });
    });
}
if(clearShippingFilterBtn) {
    clearShippingFilterBtn.addEventListener('click', () => {
        filterShippingColumn.value = '';
        filterShippingValue.value = '';
        document.querySelectorAll('#shippingScheduleTable tbody tr').forEach(row => {
            row.style.display = '';
        });
    });
}
if(clearShippingDateBtn) {
    clearShippingDateBtn.addEventListener('click', () => {
        if(shippingScheduleDateInput._flatpickr) shippingScheduleDateInput._flatpickr.clear();
    });
}

    deleteAnalisiRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('.analisi-row-selector:checked');
        if (selectedRows.length === 0) {
            await showAlert('Seleziona almeno una riga da eliminare dalla tabella Analisi.');
            return;
        }
        const confirmed = await showConfirm(`Sei sicuro di voler eliminare ${selectedRows.length} riga/e selezionata/e dalla tabella Analisi?`);
        if (confirmed) {
            selectedRows.forEach(checkbox => {
                checkbox.closest('tr').remove();
            });
            await showAlert('Riga/e eliminata/e con successo dalla tabella Analisi.');
            addLogEntry(`Eliminate ${selectedRows.length} riga/e dalla tabella Analisi.`);
        }
    });


    exportAnalisiPdfBtn.addEventListener('click', () => {
        document.body.classList.add('printing-analisi');
        window.onafterprint = () => {
            document.body.classList.remove('printing-analisi');
            window.onafterprint = null;
        };
        window.print();
    });


    logbookBtn.addEventListener('click', () => {
        logbookContainer.style.display = logbookContainer.style.display === 'none' ? 'block' : 'none';
        if (logbookContainer.style.display === 'block') {
            // Questa funzione (già presente) filtra e mostra il contenuto aggiornato
            filterAndRenderLogbook();
        }
    });

logbookBtn.addEventListener('click', () => {
    // ... codice esistente ...
});


    // ===================================================================
    // ==> 3. SOSTITUISCI L'INTERO BLOCCO DI INIZIALIZZAZIONE CON QUESTO <==
    // ===================================================================

    // --- LOGICA FILTRI ANALISI (Unificata e Corretta) ---

    const analisiSearchLottoInput = document.getElementById('searchLottoInput');
    const analisiStartDateInput = document.getElementById('analisiStartDate');
    const analisiEndDateInput = document.getElementById('analisiEndDate');
    const analisiApplyFilterBtn = document.getElementById('applyAnalisiFilterBtn');
    const analisiClearFilterBtn = document.getElementById('clearAnalisiFilterBtn');

    function saveAnalisiFiltersState() {
        if (analisiSearchLottoInput && analisiStartDateInput && analisiEndDateInput) {
            localStorage.setItem('analisi_filter_lotto', analisiSearchLottoInput.value);
            localStorage.setItem('analisi_filter_start', analisiStartDateInput.value);
            localStorage.setItem('analisi_filter_end', analisiEndDateInput.value);
        }
    }

    function loadAnalisiFiltersState() {
        const lotto = localStorage.getItem('analisi_filter_lotto');
        if (lotto && analisiSearchLottoInput) {
            analisiSearchLottoInput.value = lotto;
        }
    }

    function clearAnalisiFilters() {
        if (!analisiSearchLottoInput || !analisiStartDateInput || !analisiEndDateInput) return;
        analisiSearchLottoInput.value = '';
        const today = new Date();
        const fourteenDaysLater = new Date();
        fourteenDaysLater.setDate(today.getDate() + 14);
        if (analisiStartDateInput._flatpickr) analisiStartDateInput._flatpickr.setDate(today, true);
        if (analisiEndDateInput._flatpickr) endDateInput._flatpickr.setDate(fourteenDaysLater, true);
        applyAnalisiFilters();
    }

    if (analisiStartDateInput && analisiEndDateInput) {
    const todayForFilter = new Date();
    const fourteenDaysLaterForFilter = new Date();
    fourteenDaysLaterForFilter.setDate(todayForFilter.getDate() + 14);

    // Inizializza i calendari con il filtro live sull'evento 'onChange'
    flatpickr(analisiStartDateInput, {
        dateFormat: "d/m/Y",
        locale: "it",
        defaultDate: localStorage.getItem('analisi_filter_start') || todayForFilter,
        onChange: applyAnalisiFilters // Applica il filtro al cambio della data
    });
    flatpickr(analisiEndDateInput, {
        dateFormat: "d/m/Y",
        locale: "it",
        defaultDate: localStorage.getItem('analisi_filter_end') || fourteenDaysLaterForFilter,
        onChange: applyAnalisiFilters // Applica il filtro al cambio della data
    });

    // Attiva il filtro live mentre l'utente digita nel campo lotto
    analisiSearchLottoInput.addEventListener('input', applyAnalisiFilters);

    // Il pulsante per cancellare rimane invariato
    analisiClearFilterBtn.addEventListener('click', clearAnalisiFilters);
}


// ================================================================
// L'inizializzazione completa dell'applicazione (caricamento dati
// persistiti, preparazione delle tabelle e generazione dei grafici) è
// stata spostata in una funzione separata eseguita dopo il login.
// Questo evita che il caricamento di grosse quantità di dati blocchi
// l'interfaccia prima che l'utente possa inserire la password.
// La funzione initializeAfterLogin viene invocata dal gestore di login
// quando l'autenticazione ha successo.

// Flag che indica se l'inizializzazione dell'applicazione è già stata
// eseguita. Previene più invocazioni dello stesso blocco.
window._appInitialized = false;

function initializeAfterLogin() {
    if (window._appInitialized) return;
    window._appInitialized = true;
    try {
        gestisciCacheDatiStatici();
        // Caricamento di tutti i dati salvati in locale e da server
        caricaDatiStaticiForzatamente();
        loadAllAutoSavedData();
        loadLastModifiedTimestamp();
        loadLogbook();
        loadAnalisiFiltersState();
        // Carica e prepara i dati della tabella Medical Device dopo il login
        if (typeof loadMedicalDeviceState === 'function') {
            loadMedicalDeviceState();
        }
        // Prepara l'interfaccia (es. intestazioni della tabella analisi)
        renderAnalysisTableHeaders();
        // Renderizzazione e aggiornamento UI
        renderAnalysisTableHeaders();
        updateScrollButtons();
        updateGanttChart();
        updateWarehouseGanttChart();
        updateDailyProductionTable();
        // Effettua il rendering della tabella Medical Device una volta caricati i dati
        if (typeof renderMedicalDeviceTable === 'function') {
            try {
                renderMedicalDeviceTable();
            } catch (e) {
                console.warn('Errore nel rendering della tabella Medical Device:', e);
            }
        }
        updateAnalisiTable();
        runFullCheck();
    } catch (e) {
        console.error('Errore durante l\'inizializzazione post-login:', e);
    }
}

const exportPropostaLayoutBtn = document.getElementById('exportPropostaLayoutBtn');
    if (exportPropostaLayoutBtn) {
        exportPropostaLayoutBtn.addEventListener('click', exportPropostaLayoutPDF);
    }
    
    // Intervallo di aggiornamento automatico dei dati impostato a 5 minuti
    // (300.000 ms). Questo consente a tutti gli utenti di ricevere in modo
    // regolare le modifiche effettuate da altri sulle tabelle. Per cambiare
    // l'intervallo di sincronizzazione è sufficiente modificare il valore.
    setInterval(loadDataFromServer, 300000);
    window.addEventListener('resize', updateStickyPositions);
    const mainObserver = new MutationObserver(updateStickyPositions);
    if (stickyControlsWrapper) {
        mainObserver.observe(stickyControlsWrapper, { childList: true, subtree: true, attributes: true });
    }
    updateStickyPositions();
    // Prima di rendere la tabella ridimensionabile, verifichiamo che la funzione sia disponibile.
    if (typeof window !== 'undefined' && typeof window.makeTableResizable === 'function') {
        window.makeTableResizable(document.getElementById('shippingScheduleTable'));
    }
// ===================================================================
    // ==> 2. BLOCCO NUOVO DA AGGIUNGERE <==
    // Logica per il pulsante di ricarica manuale
    // ===================================================================
    if (manualRefreshBtn) {
        manualRefreshBtn.addEventListener('click', async () => {
            if (await showConfirm("Vuoi forzare l'aggiornamento dei dati dal server? Le modifiche non salvate andranno perse.")) {
                addLogEntry('Ricarica manuale dei dati dal server avviata dall\'utente.');
                await loadDataFromServer();
                showAlert('Ricarica completata con successo!');
            }
        });
    }

// ========================================================================
// ==> BLOCCO EVENT LISTENERS PER TABELLA ARRIVI (COMPLETO E CORRETTO)
// ========================================================================
    const importLayoutBtn = document.getElementById('importLayoutBtn');
    if (importLayoutBtn) {
        // Gestore aggiornato: crea un input file autonomo in modo da non
        // interferire con altri import e consente al magazziniere di
        // selezionare file layout.  Dopo l'import aggiorna la data
        // dell'ultimo caricamento e verifica eventuali ADR.
        importLayoutBtn.addEventListener('click', () => {
            const layoutInput = document.createElement('input');
            layoutInput.type = 'file';
            layoutInput.accept = '.xls,.xlsx,.csv';
            layoutInput.style.display = 'none';
            document.body.appendChild(layoutInput);
            layoutInput.addEventListener('change', async (e) => {
                const file = e.target.files[0];
                if (!file) {
                    document.body.removeChild(layoutInput);
                    return;
                }
                try {
                    await processLayoutFile(file);
                    // Salva la data di importazione per visualizzarla accanto al bottone
                    if (typeof formatDateTimeForDisplay === 'function') {
                        const nowStr = formatDateTimeForDisplay(new Date());
                        // Usa una chiave uniforme senza underscore
                        localStorage.setItem('lastImportLayout', nowStr);
                    } else {
                        localStorage.setItem('lastImportLayout', Date.now().toString());
                    }
                    if (typeof updateImportTimestamps === 'function') {
                        updateImportTimestamps();
                    }
                    // Controlla nuovamente eventuali spedizioni ADR dopo aver importato il layout
                    if (typeof checkAndNotifyADR === 'function') {
                        checkAndNotifyADR();
                    }
                } catch (err) {
                    console.error('Errore durante l\'importazione del file Layout:', err);
                } finally {
                    // Rimuovi l'input dal DOM
                    document.body.removeChild(layoutInput);
                }
            });
            layoutInput.click();
        });
    }

if (addArrivalRowBtn) {
    addArrivalRowBtn.addEventListener('click', () => {
        arrivalScheduleTableBody.appendChild(createArrivalScheduleRow());
        // Rende nuovamente ordinabile la tabella arrivi
        if (typeof makeTableSortable === 'function') {
            makeTableSortable(document.getElementById('arrivalScheduleTable'));
        }
        autoSaveAllData();
    });
}
if (duplicateArrivalRowBtn) {
    duplicateArrivalRowBtn.addEventListener('click', () => {
        const selectedRows = document.querySelectorAll('#arrivalScheduleTable .arrival-row-selector:checked');
        selectedRows.forEach(checkbox => {
            const originalRow = checkbox.closest('tr');
            const rowData = getArrivalScheduleRowData(originalRow);
            originalRow.after(createArrivalScheduleRow(rowData));
        });
        updateWarehouseGanttChart();
        // Rende di nuovo ordinabile la tabella arrivi
        if (typeof makeTableSortable === 'function') {
            makeTableSortable(document.getElementById('arrivalScheduleTable'));
        }
        autoSaveAllData();
    });
}
if (deleteArrivalRowBtn) {
    deleteArrivalRowBtn.addEventListener('click', async () => {
        const selectedRows = document.querySelectorAll('#arrivalScheduleTable .arrival-row-selector:checked');
        if (selectedRows.length > 0 && await showConfirm(`Eliminare ${selectedRows.length} righe di arrivi?`)) {
            selectedRows.forEach(checkbox => checkbox.closest('tr').remove());
            updateWarehouseGanttChart();
            // Aggiorna l'ordinabilità della tabella dopo eliminazione
            if (typeof makeTableSortable === 'function') {
                makeTableSortable(document.getElementById('arrivalScheduleTable'));
            }
            autoSaveAllData();
        }
    });
}

// Gestione eliminazione righe dalla tabella "Merce in Quarantena"
const deleteQuarantineRowBtn = document.getElementById('deleteQuarantineRowBtn');
if (deleteQuarantineRowBtn) {
    deleteQuarantineRowBtn.addEventListener('click', async () => {
        // Seleziona tutte le checkbox selezionate nella tabella quarantena
        const selected = document.querySelectorAll('#quarantineTable .quarantine-row-selector:checked');
        if (selected.length > 0) {
            // Mostra una conferma con il numero di righe da cancellare
            if (await showConfirm(`Eliminare ${selected.length} righe dalla quarantena?`)) {
                selected.forEach(checkbox => {
                    const tr = checkbox.closest('tr');
                    if (tr) tr.remove();
                });
                // Dopo la cancellazione, salva i dati
                autoSaveAllData();
                if (typeof saveDataToServer === 'function') {
                    try {
                        await saveDataToServer();
                    } catch (e) {
                        console.warn('Errore nel salvataggio dati dopo eliminazione quarantena:', e);
                    }
                }
            }
        }
    });
}

function applyArrivalFilter() {
    if (!filterArrivalColumn || !filterArrivalValue || !arrivalStartDateInput || !arrivalEndDateInput) return;

    const filterCol = filterArrivalColumn.value;
    const filterVal = filterArrivalValue.value.trim().toLowerCase();
    const startDate = flatpickr.parseDate(arrivalStartDateInput.value, "d/m/Y");
    const endDate = flatpickr.parseDate(arrivalEndDateInput.value, "d/m/Y");
    if (endDate) endDate.setHours(23, 59, 59, 999);

    document.querySelectorAll('#arrivalScheduleTable tbody tr').forEach(row => {
        let showRow = true;
        const rowData = getArrivalScheduleRowData(row);

        if (filterCol && filterVal && !String(rowData[filterCol] || '').toLowerCase().includes(filterVal)) {
            showRow = false;
        }

        if (showRow && (startDate || endDate)) {
            const rowDateParts = rowData.dataConsegna.split('/');
            if (rowDateParts.length === 3) {
                const rowDate = new Date(parseInt(rowDateParts[2]), parseInt(rowDateParts[1]) - 1, parseInt(rowDateParts[0]));
                if ((startDate && rowDate < startDate) || (endDate && rowDate > endDate)) {
                    showRow = false;
                }
            } else {
                showRow = false; 
            }
        }
        row.style.display = showRow ? '' : 'none';
    });
}

// Attivazione e inizializzazione dei filtri per la tabella arrivi
if(applyArrivalFilterBtn) applyArrivalFilterBtn.addEventListener('click', applyArrivalFilter);
if(clearArrivalFilterBtn) clearArrivalFilterBtn.addEventListener('click', () => {
    filterArrivalColumn.value = '';
    filterArrivalValue.value = '';
    if (arrivalStartDateInput._flatpickr) arrivalStartDateInput._flatpickr.clear();
    if (arrivalEndDateInput._flatpickr) arrivalEndDateInput._flatpickr.clear();
    applyArrivalFilter();
});
if(clearArrivalDateBtn) {
    clearArrivalDateBtn.addEventListener('click', () => {
        if (arrivalStartDateInput._flatpickr) arrivalStartDateInput._flatpickr.clear();
        if (arrivalEndDateInput._flatpickr) arrivalEndDateInput._flatpickr.clear();
        applyArrivalFilter();
    });
}

// --- VERSIONE AGGIORNATA PER ARRIVI ---

// Attiva i filtri in tempo reale
if(applyArrivalFilterBtn) applyArrivalFilterBtn.style.display = 'none'; // Nasconde il vecchio pulsante se esiste ancora
if(filterArrivalColumn) filterArrivalColumn.addEventListener('change', applyArrivalFilter);
if(filterArrivalValue) filterArrivalValue.addEventListener('input', applyArrivalFilter);

// Inizializzazione dei calendari con filtro "live"
if (arrivalStartDateInput && arrivalEndDateInput) {
    const oggi = new Date();
    const settimanaProssima = new Date();
    settimanaProssima.setDate(oggi.getDate() + 7);

    flatpickr(arrivalStartDateInput, {
        dateFormat: "d/m/Y",
        locale: "it",
        defaultDate: oggi,
        onChange: applyArrivalFilter // Filtra al cambio della data
    });
    flatpickr(arrivalEndDateInput, {
        dateFormat: "d/m/Y",
        locale: "it",
        defaultDate: settimanaProssima,
        onChange: applyArrivalFilter // Filtra al cambio della data
    });
    
    setTimeout(applyArrivalFilter, 500);
}
// ========================================================================
    // ==> BLOCCO DI ATTIVAZIONE PER LA NUOVA TABELLA <==
    // ========================================================================
    if (medicalDeviceStartDateInput && medicalDeviceEndDateInput) {
        const oggi = new Date();
        const nextWeek = new Date();
        nextWeek.setDate(oggi.getDate() + 7);

        flatpickr(medicalDeviceStartDateInput, { dateFormat: "d/m/Y", locale: "it", defaultDate: oggi, onChange: updateMedicalDeviceProductionTable });
        flatpickr(medicalDeviceEndDateInput, { dateFormat: "d/m/Y", locale: "it", defaultDate: nextWeek, onChange: updateMedicalDeviceProductionTable });

        filterMedicalDeviceCodice.addEventListener('input', updateMedicalDeviceProductionTable);
        filterMedicalDeviceDescrizione.addEventListener('input', updateMedicalDeviceProductionTable);
        filterMedicalDeviceCliente.addEventListener('input', updateMedicalDeviceProductionTable);
        // Listener anche per i nuovi filtri data e lotto se presenti
        if (filterMedicalDeviceData) filterMedicalDeviceData.addEventListener('input', updateMedicalDeviceProductionTable);
        if (filterMedicalDeviceLotto) filterMedicalDeviceLotto.addEventListener('input', updateMedicalDeviceProductionTable);
        // Nuovi listener per filtri data e lotto
        if (filterMedicalDeviceData) filterMedicalDeviceData.addEventListener('input', updateMedicalDeviceProductionTable);
        if (filterMedicalDeviceLotto) filterMedicalDeviceLotto.addEventListener('input', updateMedicalDeviceProductionTable);

        clearMedicalDeviceDateBtn.addEventListener('click', () => {
            medicalDeviceStartDateInput._flatpickr.clear();
            medicalDeviceEndDateInput._flatpickr.clear();
            updateMedicalDeviceProductionTable();
        });
        clearMedicalDeviceFiltersBtn.addEventListener('click', () => {
            filterMedicalDeviceCodice.value = '';
            filterMedicalDeviceDescrizione.value = '';
            filterMedicalDeviceCliente.value = '';
            if (filterMedicalDeviceData) filterMedicalDeviceData.value = '';
            if (filterMedicalDeviceLotto) filterMedicalDeviceLotto.value = '';
            updateMedicalDeviceProductionTable();
        });
    }

    if (addMedicalDeviceRowBtn) {
        addMedicalDeviceRowBtn.addEventListener('click', () => {
            medicalDeviceTableBody.appendChild(createMedicalDeviceRow({}, true));
        });
    }

// ========================================================================
    // ==> BLOCCO DI ATTIVAZIONE PER LA NUOVA TABELLA MEDICAL DEVICE <==
    // ========================================================================
    if (medicalDeviceStartDateInput && medicalDeviceEndDateInput) {
        const oggi = new Date();
        const nextWeek = new Date();
        nextWeek.setDate(oggi.getDate() + 7);

        flatpickr(medicalDeviceStartDateInput, { dateFormat: "d/m/Y", locale: "it", defaultDate: oggi, onChange: updateMedicalDeviceProductionTable });
        flatpickr(medicalDeviceEndDateInput, { dateFormat: "d/m/Y", locale: "it", defaultDate: nextWeek, onChange: updateMedicalDeviceProductionTable });

        filterMedicalDeviceCodice.addEventListener('input', updateMedicalDeviceProductionTable);
        filterMedicalDeviceDescrizione.addEventListener('input', updateMedicalDeviceProductionTable);
        filterMedicalDeviceCliente.addEventListener('input', updateMedicalDeviceProductionTable);

        clearMedicalDeviceDateBtn.addEventListener('click', () => {
            medicalDeviceStartDateInput._flatpickr.clear();
            medicalDeviceEndDateInput._flatpickr.clear();
            updateMedicalDeviceProductionTable();
        });
        clearMedicalDeviceFiltersBtn.addEventListener('click', () => {
            filterMedicalDeviceCodice.value = '';
            filterMedicalDeviceDescrizione.value = '';
            filterMedicalDeviceCliente.value = '';
            if (filterMedicalDeviceData) filterMedicalDeviceData.value = '';
            if (filterMedicalDeviceLotto) filterMedicalDeviceLotto.value = '';
            updateMedicalDeviceProductionTable();
        });
    }

    if (addMedicalDeviceRowBtn) {
        addMedicalDeviceRowBtn.addEventListener('click', () => {
            medicalDeviceTableBody.appendChild(createMedicalDeviceRow({}, true));
        });
    }

// Aggiungi questo listener alla fine del tuo script `DOMContentLoaded`
    // Aggiungi questo listener alla fine del tuo script `DOMContentLoaded`
document.addEventListener('click', (e) => {
    if (genericTooltip && genericTooltip.classList.contains('visible')) {
        // Controlla se il click è avvenuto fuori dal tooltip
        const clickedTask = e.target.closest('.gantt-task');
        // La riga successiva non è più necessaria per la logica di chiusura, ma la lasciamo per non alterare la struttura
        const isShippingTask = clickedTask ? clickedTask.classList.contains('shipping-task') : false;

        // VERSIONE CORRETTA
        if (!genericTooltip.contains(e.target)) {
            hideGenericTooltip();
        }
    }
});

// ------------------------------------------------------------------------
//  Funzione di utilità: abbreviazione dei nomi degli operatori per la stampa
//  Questa funzione prende un nome completo (cognome e uno o più nomi) e
//  restituisce il cognome seguito dalle iniziali dei nomi, ciascuna con un
//  punto.  Esempi:
//    "Motta Diego"    -> "Motta D."
//    "Motta Diego Luca" -> "Motta D. L."
//    "Murriani Giuseppe" -> "Murriani G."
//  Se il nome contiene solo il cognome, viene restituito senza modifiche.
function abbreviateOperatorName(fullName) {
    const trimmed = String(fullName || '').trim();
    if (!trimmed) return '';
    const parts = trimmed.split(/\s+/);
    if (parts.length === 1) {
        return parts[0];
    }
    const surname = parts[0];
    const initials = parts.slice(1).map(p => {
        const c = p.charAt(0);
        return c ? c.toUpperCase() + '.' : '';
    }).join(' ');
    return initials ? `${surname} ${initials}` : surname;
}
 });


    </script>
    <script>
    /* -------------------------------------------------------------------
     * Packing List functionality
     * These functions allow the user to generate a printable packing list
     * from the list of shipping orders scheduled in the next two weeks.
     * ------------------------------------------------------------------- */
    let packingListData = {};
    function escapeHtml(text) {
        return String(text).replace(/&/g, '&amp;')
                          .replace(/</g, '&lt;')
                          .replace(/>/g, '&gt;')
                          .replace(/"/g, '&quot;')
                          .replace(/'/g, '&#39;');
    }
    function openPackingListModal() {
        const modal = document.getElementById('packingListModal');
        const list = document.getElementById('packingListItems');
        if (!modal || !list) return;
        // Gestisce la visibilità dei pulsanti di creazione e chiusura del modale Packing List
        {
            const createBtn = document.getElementById('packingListCreateBtn');
            const closeBtn = document.getElementById('packingListCloseBtn');
            // Il pulsante di chiusura deve sempre essere visibile per permettere l'uscita dal modale
            if (closeBtn) {
                closeBtn.style.display = 'inline-flex';
                closeBtn.disabled = false;
                closeBtn.style.pointerEvents = 'auto';
            }
            if (createBtn) {
                /*
                 * In questa versione mostriamo sempre il pulsante "Crea Packing List"
                 * per i livelli abilitati tramite applyPermissions.  La visibilità
                 * del bottone principale viene gestita a monte in base al ruolo,
                 * pertanto qui non è necessario nascondere il pulsante.  Ci
                 * limitiamo a renderlo interattivo.
                 */
                createBtn.style.display = 'inline-flex';
                createBtn.disabled = false;
                createBtn.style.pointerEvents = 'auto';
            }
        }
        list.innerHTML = '';
        packingListData = {};
        // Recupera i dati di spedizione. In alcuni casi la tabella del programma
        // di spedizione potrebbe non essere stata ancora renderizzata oppure
        // getAllShippingData() può restituire un array vuoto. In tal caso
        // effettua un fallback caricando i dati memorizzati nel localStorage
        // sotto la chiave "shipping_schedule_data_autosave". Questo valore
        // viene salvato automaticamente quando l’utente apporta modifiche al
        // programma di spedizione, permettendo di ripopolare il modale
        // Packing List anche dopo un reload.
        let shippingData = [];
        if (typeof getAllShippingData === 'function') {
            try {
                shippingData = getAllShippingData() || [];
            } catch (e) {
                shippingData = [];
            }
        }
        // Se non sono presenti dati nella tabella (es. appena caricata la
        // pagina o senza import), prova a leggere i dati salvati
        // automaticamente dal localStorage.  Il parsing è protetto da
        // try/catch per evitare errori in caso di JSON malformato.
        if (!Array.isArray(shippingData) || shippingData.length === 0) {
            try {
                const saved = localStorage.getItem('shipping_schedule_data_autosave');
                if (saved) {
                    const parsed = JSON.parse(saved);
                    if (Array.isArray(parsed)) {
                        shippingData = parsed;
                    }
                }
            } catch (err) {
                shippingData = [];
            }
        }
        const today = new Date();
        today.setHours(0,0,0,0);
        const twoWeeks = new Date(today);
        twoWeeks.setDate(today.getDate() + 13);
        twoWeeks.setHours(23,59,59,999);
        // Oggetto per raggruppare gli ordini per data e OV
        const packingByDate = {};
        shippingData.forEach(item => {
            const parts = String(item.dataConsegna || '').split('/');
            if (parts.length !== 3) return;
            const d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
            d.setHours(0,0,0,0);
            if (d < today || d > twoWeeks) return;
            const dateKey = String(item.dataConsegna || '').trim();
            const ov = String(item.ov || '').trim();
            if (!ov) return;
            // Popola la struttura originale per compatibilità con la stampa
            if (!packingListData[ov]) packingListData[ov] = [];
            packingListData[ov].push(item);
            // Popola la struttura per gruppo per data/ov
            if (!packingByDate[dateKey]) packingByDate[dateKey] = {};
            if (!packingByDate[dateKey][ov]) packingByDate[dateKey][ov] = [];
            packingByDate[dateKey][ov].push(item);
        });
        // Ordina le date (convertendo gg/mm/aaaa in un oggetto Date per ordinamento)
        const dates = Object.keys(packingByDate).sort((a,b) => {
            const [d1, m1, y1] = a.split('/');
            const [d2, m2, y2] = b.split('/');
            const dateA = new Date(parseInt(y1,10), parseInt(m1,10)-1, parseInt(d1,10));
            const dateB = new Date(parseInt(y2,10), parseInt(m2,10)-1, parseInt(d2,10));
            return dateA - dateB;
        });
        dates.forEach(dateKey => {
            // Inserisci un'intestazione di data senza checkbox
            const dateHeaderLi = document.createElement('li');
            dateHeaderLi.className = 'date-header';
            dateHeaderLi.textContent = dateKey;
            list.appendChild(dateHeaderLi);
            const ovMap = packingByDate[dateKey] || {};
            const ovKeys = Object.keys(ovMap).sort((a,b) => a.localeCompare(b));
            ovKeys.forEach(ov => {
                const group = ovMap[ov];
                const parentLi = document.createElement('li');
                parentLi.className = 'packing-list-item';
                const expandBtn = document.createElement('span');
                expandBtn.textContent = '+';
                expandBtn.style.cursor = 'pointer';
                expandBtn.style.marginRight = '6px';
                expandBtn.style.fontWeight = 'bold';
                const parentCheckbox = document.createElement('input');
                parentCheckbox.type = 'checkbox';
                parentCheckbox.dataset.ov = ov;
                parentCheckbox.style.marginRight = '8px';
                const label = document.createElement('label');
                const first = group[0];
                // Mostra la data di consegna tra parentesi, se disponibile (è uguale a dateKey)
                label.textContent = `OV ${ov} - ${first.ragioneSociale || ''}${dateKey ? ` (${dateKey})` : ''}`;
                const childUl = document.createElement('ul');
                childUl.style.display = 'none';
                group.forEach((item, index) => {
                    const childLi = document.createElement('li');
                    const cb = document.createElement('input');
                    cb.type = 'checkbox';
                    cb.dataset.ov = ov;
                    cb.dataset.index = String(index);
                    cb.style.marginRight = '6px';
                    const desc = `${item.codiceArticolo || ''} - ${item.descrizioneArticolo || ''} (${item.quantita || ''} ${item.um || ''})`;
                    const span = document.createElement('span');
                    span.textContent = desc;
                    childLi.appendChild(cb);
                    childLi.appendChild(span);
                    childUl.appendChild(childLi);
                });
                expandBtn.addEventListener('click', () => {
                    if (childUl.style.display === 'none') {
                        childUl.style.display = 'block';
                        expandBtn.textContent = '-';
                    } else {
                        childUl.style.display = 'none';
                        expandBtn.textContent = '+';
                    }
                });
                parentCheckbox.addEventListener('change', () => {
                    const checked = parentCheckbox.checked;
                    const childBoxes = childUl.querySelectorAll('input[type="checkbox"]');
                    childBoxes.forEach(cb => {
                        cb.checked = checked;
                    });
                    parentCheckbox.indeterminate = false;
                });
                childUl.addEventListener('change', () => {
                    const childBoxes = childUl.querySelectorAll('input[type="checkbox"]');
                    let checkedCount = 0;
                    childBoxes.forEach(cb => {
                        if (cb.checked) checkedCount++;
                    });
                    if (checkedCount === 0) {
                        parentCheckbox.checked = false;
                        parentCheckbox.indeterminate = false;
                    } else if (checkedCount === childBoxes.length) {
                        parentCheckbox.checked = true;
                        parentCheckbox.indeterminate = false;
                    } else {
                        parentCheckbox.checked = false;
                        parentCheckbox.indeterminate = true;
                    }
                });
                parentLi.appendChild(expandBtn);
                parentLi.appendChild(parentCheckbox);
                parentLi.appendChild(label);
                parentLi.appendChild(childUl);
                list.appendChild(parentLi);
            });
        });
        modal.style.display = 'flex';
    }
    function closePackingListModal() {
        const modal = document.getElementById('packingListModal');
        if (modal) modal.style.display = 'none';
    }
    function buildPackingList() {
        const modal = document.getElementById('packingListModal');
        const selected = [];
        const groups = document.querySelectorAll('#packingListItems > li');
        groups.forEach(li => {
            const parentCb = li.querySelector('input[type="checkbox"]');
            if (!parentCb) return;
            const ov = parentCb.dataset.ov;
            const childCbs = li.querySelectorAll('ul input[type="checkbox"]');
            const selectedItems = [];
            childCbs.forEach(cb => {
                if (cb.checked) {
                    const idx = parseInt(cb.dataset.index, 10);
                    const item = packingListData[ov] ? packingListData[ov][idx] : null;
                    if (item) selectedItems.push(item);
                }
            });
            if (selectedItems.length > 0) {
                selected.push({ ov: ov, items: selectedItems });
            } else if (parentCb.checked) {
                selected.push({ ov: ov, items: packingListData[ov] });
            }
        });
        if (selected.length === 0) {
            alert('Seleziona almeno un ordine di vendita da includere.');
            return;
        }
        let html = `<html><head><title>Packing List</title><style>\n` +
            `body { font-family: Arial, sans-serif; margin: 20px; }\n` +
            `h1 { margin-top: 0; }\n` +
            `h2 { margin-top: 30px; }\n` +
            `table { width: 100%; border-collapse: collapse; margin-top: 10px; }\n` +
            `th, td { border: 1px solid #ccc; padding: 5px; font-size: 0.85em; }\n` +
            `th { background-color: #f5f5f5; }\n` +
            `.header-table { width: 100%; margin-bottom: 10px; }\n` +
            `.header-table td { border: none; padding: 2px 4px; }\n` +
            `</style></head><body>`;
        html += `<h1>Packing List</h1>`;
        const currentDate = new Date().toLocaleDateString('it-IT');
        html += `<p>Data: ${currentDate}</p>`;
        selected.forEach(group => {
            const first = group.items[0];
            const shippingDate = first && first.dataConsegna ? escapeHtml(first.dataConsegna) : '';
            html += `<h2>Ordine di Vendita: ${escapeHtml(group.ov)}${shippingDate ? ' - Data di Spedizione: ' + shippingDate : ''}</h2>`;
            html += `<table class="header-table"><tr><td><strong>Ragione Sociale:</strong> ${escapeHtml(first.ragioneSociale || '')}</td></tr>`;
            html += `<tr><td><strong>Indirizzo:</strong> ${escapeHtml(first.indirizzo || '')}</td></tr>`;
            html += `<tr><td><strong>Località:</strong> ${escapeHtml(first.cap || '')} ${escapeHtml(first.citta || '')} (${escapeHtml(first.provincia || '')})</td></tr>`;
            html += `</table>`;
            html += `<table><thead><tr><th>Codice</th><th>Descrizione</th><th>Quantità</th><th>Lotto</th><th>Data Produzione</th><th>Data Scadenza</th><th>Note</th></tr></thead><tbody>`;
            group.items.forEach(item => {
                html += `<tr><td>${escapeHtml(item.codiceArticolo || '')}</td>`;
                html += `<td>${escapeHtml(item.descrizioneArticolo || '')}</td>`;
                html += `<td>${escapeHtml(item.quantita || '')} ${escapeHtml(item.um || '')}</td>`;
                html += `<td></td><td></td><td></td><td></td></tr>`;
            });
            html += `<tr><td colspan="7" style="padding-top:10px;"><strong>Numero colli:</strong> ______ &nbsp;&nbsp; <strong>Numero pallet:</strong> ______ &nbsp;&nbsp; <strong>Peso totale:</strong> ______</td></tr>`;
            html += `</tbody></table>`;
            html += `<hr style="margin-top:20px;">`;
        });
        html += `</body></html>`;
        const printWin = window.open('', '_blank');
        printWin.document.write(html);
        printWin.document.close();
        printWin.focus();
        setTimeout(() => {
            printWin.print();
        }, 500);
        if (modal) modal.style.display = 'none';
    }

    /* ====================================================================
     * OVERRIDE: buildPackingList
     * Questa versione sostituisce la funzione esistente per generare la
     * packing list con un layout orizzontale, colonne ampliate e dati
     * provenienti da tutte le tabelle disponibili (OPI, produzione
     * medicale, DeviceRef).  Vengono incluse colonne per OP, lotto,
     * pezzi reali, aghi/valva, siringhe per scatola, pesi e ADR.
     * Include riga logistica distanziata e riga note su tre righe.
     * Evita la divisione delle tabelle su più pagine tramite
     * page-break-avoid e page-break-after.
     * ==================================================================== */
    buildPackingList = function() {
        const modal = document.getElementById('packingListModal');
        const selected = [];
        const groups = document.querySelectorAll('#packingListItems > li');
        groups.forEach(li => {
            const parentCb = li.querySelector('input[type="checkbox"]');
            if (!parentCb) return;
            const ov = parentCb.dataset.ov;
            const childCbs = li.querySelectorAll('ul input[type="checkbox"]');
            const selectedItems = [];
            childCbs.forEach(cb => {
                if (cb.checked) {
                    const idx = parseInt(cb.dataset.index, 10);
                    const item = packingListData[ov] ? packingListData[ov][idx] : null;
                    if (item) selectedItems.push(item);
                }
            });
            if (selectedItems.length > 0) {
                selected.push({ ov: ov, items: selectedItems });
            } else if (parentCb.checked) {
                selected.push({ ov: ov, items: packingListData[ov] });
            }
        });
        if (selected.length === 0) {
            alert('Seleziona almeno un ordine di vendita da includere.');
            return;
        }
        // Raccoglie i dati dalle tabelle OPI, produzione medicale e DeviceRef
        let opiData = [];
        try {
            if (typeof getOpiMonitorData === 'function') {
                opiData = getOpiMonitorData() || [];
            } else {
                const opiStr = localStorage.getItem('opi_monitor_data');
                opiData = opiStr ? JSON.parse(opiStr) : [];
            }
        } catch (e) {
            opiData = [];
        }
        let medProdData = [];
        try {
            if (typeof getMedicalProductionData === 'function') {
                medProdData = getMedicalProductionData() || [];
            } else {
                const mpStr = localStorage.getItem('medicalProductionData');
                medProdData = mpStr ? JSON.parse(mpStr) : [];
            }
        } catch (e) {
            medProdData = [];
        }
        let deviceRefs = [];
        try {
            if (typeof getDeviceRefData === 'function') {
                deviceRefs = getDeviceRefData() || [];
            } else {
                const drStr = localStorage.getItem('deviceRefData');
                deviceRefs = drStr ? JSON.parse(drStr) : [];
            }
        } catch (e) {
            deviceRefs = [];
        }
        const adrSet = (window.adrCodes instanceof Set) ? window.adrCodes : new Set();
        let html = `<html><head><title>Packing List</title><style>\n` +
                   `@media print { @page { size: A4 landscape; margin: 10mm; } }\n` +
                   `body { font-family: 'Arial','Helvetica',sans-serif; margin: 10mm; }\n` +
                   `h1 { margin: 0 0 6px 0; }\n` +
                   `h2 { margin: 18px 0 6px 0; }\n` +
                   `table { width: 100%; border-collapse: collapse; margin-top: 6px; page-break-inside: avoid; }\n` +
                   `th, td { border: 1px solid #aaa; padding: 6px 4px; font-size: 9pt; text-align: center; vertical-align: middle; }\n` +
                   `th { background-color: #f0f0f0; font-weight: bold; }\n` +
                   `.header-table { width: 100%; margin-bottom: 4px; }\n` +
                   `.header-table td { border: none; padding: 2px 4px; font-size: 9pt; text-align: left; }\n` +
                   `.logistic-row td { border: none; padding: 8px 0; font-size: 9pt; text-align: left; }\n` +
                   `</style></head><body>`;
        html += `<h1>Packing List</h1>`;
        const currentDate = new Date().toLocaleDateString('it-IT');
        html += `<p>Data: ${currentDate}</p>`;
        selected.forEach((group, idx) => {
            const first = group.items[0];
            const shippingDate2 = first && first.dataConsegna ? escapeHtml(first.dataConsegna) : '';
            html += `<h2>Ordine di Vendita: ${escapeHtml(group.ov)}${shippingDate2 ? ' - Data di Spedizione: ' + shippingDate2 : ''}</h2>`;
            html += `<table class="header-table"><tr><td><strong>Ragione Sociale:</strong> ${escapeHtml(first.ragioneSociale || '')}</td></tr>`;
            html += `<tr><td><strong>Indirizzo:</strong> ${escapeHtml(first.indirizzo || '')}</td></tr>`;
            html += `<tr><td><strong>Località:</strong> ${escapeHtml(first.cap || '')} ${escapeHtml(first.citta || '')} (${escapeHtml(first.provincia || '')})</td></tr></table>`;
            // Tabella principale della packing list: includi colonne aggiuntive per
            // la quantità reale di siringhe e il numero di pezzi per scatolone.
            html += `<table><thead><tr><th>OP</th><th>Codice</th><th>Descrizione</th><th>Lotto</th><th>Q.tà Ordine</th><th>UM</th><th>Produzione</th><th>Scadenza</th><th>Quantità reale</th><th>Pezzi/scatolone</th><th>Scatoloni (ip.)</th><th>Peso scatola</th><th>Peso scatolone</th><th>ADR</th></tr></thead><tbody>`;
            group.items.forEach(item => {
                const code = (item.codiceArticolo || '').toString().trim();
                const desc = item.descrizioneArticolo || '';
                const qty = item.quantita || '';
                const um = item.um || '';
                let op = '';
                let lotto = '';
                let prodDate = '';
                let scadDate = '';
                let realPieces = '';
                let aghiPerValva = '';
                let sirPerScatola = '';
                let pesoScatola = '';
                let pesoScatolone = '';
                // Nuovi campi MD: quantità reale e pezzi per scatolone
                let mdQty = '';
                let perBoxVal = '';
                let mdBoxes = '';
                if (Array.isArray(opiData) && opiData.length > 0) {
                    const match = opiData.find(opi => String(opi.ov || '').trim().toUpperCase() === String(group.ov || '').trim().toUpperCase() && String(opi.codice || '').trim().toUpperCase() === code.toUpperCase());
                    if (match) {
                        op = match.op || '';
                        lotto = match.lotto || '';
                        prodDate = match.dataProd || '';
                        scadDate = match.scadenza || '';
                    }
                }
                if (Array.isArray(medProdData) && medProdData.length > 0) {
                    const mMatch = medProdData.find(m => {
                        const codeMatch = String(m.codice || '').trim().toUpperCase() === code.toUpperCase();
                        if (!codeMatch) return false;
                        if (lotto) return String(m.lotto || '').trim().toUpperCase() === lotto.toUpperCase();
                        return true;
                    });
                    if (mMatch && mMatch.quantita !== undefined && mMatch.quantita !== null && mMatch.quantita !== '') {
                        realPieces = String(mMatch.quantita);
                    }
                }
                if (Array.isArray(deviceRefs) && deviceRefs.length > 0) {
                    const refMatch = deviceRefs.find(ref => String(ref.codice || '').trim().toUpperCase() === code.toUpperCase());
                    if (refMatch) {
                        aghiPerValva = refMatch.aghiPerValva || '';
                        // Determina i pezzi per scatolone in base al campo più affidabile
                        const rawPerBox = (refMatch.pezziPerScatolone !== undefined ? refMatch.pezziPerScatolone : undefined) || refMatch.siringhePerScatola2 || refMatch.siringhePerScatola;
                        if (rawPerBox != null && rawPerBox !== '') {
                            perBoxVal = String(rawPerBox);
                            // normalizza per i calcoli numerici (rimuove punti come separatori migliaia, virgola come decimale)
                            const normalized = perBoxVal.replace(/\./g, '').replace(',', '.');
                            const parsedPB = parseFloat(normalized);
                            if (!isNaN(parsedPB) && parsedPB > 0) {
                                // Calcola il numero teorico di scatoloni se disponibile una quantità reale
                                const qtyForMd = parseFloat(realPieces);
                                if (!isNaN(qtyForMd) && qtyForMd > 0) {
                                    mdBoxes = Math.ceil(qtyForMd / parsedPB).toString();
                                }
                            }
                        }
                        sirPerScatola = refMatch.siringhePerScatola || refMatch.siringhePerScatola2 || '';
                        pesoScatola = refMatch.pesoScatola || '';
                        pesoScatolone = refMatch.pesoScatolone || '';
                    }
                }

                // Se realPieces non è disponibile o zero, tenta di recuperare dal packingListData (quantitaReale o pezziReali)
                if (!(parseFloat(realPieces) > 0)) {
                    try {
                        const ovKey = String(group.ov || '').trim();
                        const list = window.packingListData && window.packingListData[ovKey];
                        if (Array.isArray(list)) {
                            const itemPL = list.find(it => String(it.codiceArticolo || '').trim().toUpperCase() === code.toUpperCase());
                            const val = itemPL && (itemPL.quantitaReale || itemPL.pezziReali);
                            if (val) {
                                const parsed = parseFloat(String(val).replace(/\./g, '').replace(',', '.'));
                                if (!isNaN(parsed) && parsed > 0) realPieces = String(parsed);
                            }
                        }
                    } catch (_) {}
                }
                // Assegna mdQty come numero reale (realPieces) per chiarezza nel rendering
                mdQty = realPieces;

                const isAdr = adrSet.has(code.toUpperCase()) ? 'ADR' : '';
                html += `<tr><td>${escapeHtml(op)}</td><td>${escapeHtml(code)}</td><td>${escapeHtml(desc)}</td><td>${escapeHtml(lotto)}</td><td>${escapeHtml(qty)}</td><td>${escapeHtml(um)}</td><td>${escapeHtml(prodDate)}</td><td>${escapeHtml(scadDate)}</td><td>${escapeHtml(mdQty)}</td><td>${escapeHtml(perBoxVal)}</td><td>${escapeHtml(mdBoxes)}</td><td>${escapeHtml(pesoScatola)}</td><td>${escapeHtml(pesoScatolone)}</td><td>${isAdr}</td></tr>`;
            });
            // Spaziatura prima della riga logistica
            html += `<tr><td colspan="14" style="height:24px; border:none;"></td></tr>`;
            // Riga logistica con più spazio per i campi da compilare
            html += `<tr class="logistic-row"><td colspan="14"><strong>Numero colli:</strong> _______ &nbsp;&nbsp;&nbsp; <strong>Numero pallet:</strong> ___________ &nbsp;&nbsp;&nbsp; <strong>Peso netto totale:</strong> _________ kg &nbsp;&nbsp;&nbsp; <strong>Peso lordo totale:</strong> __________ kg &nbsp;&nbsp;&nbsp; <strong>Volume totale:</strong> ___________ m³</td></tr>`;
            // Sezione note commerciali
            html += `<tr class="logistic-row"><td colspan="14"><strong>Note commerciali:</strong><br><span style="display:inline-block; width:100%; border-bottom:1px solid #aaa; height:12px;"></span><br><span style="display:inline-block; width:100%; border-bottom:1px solid #aaa; height:12px;"></span><br><span style="display:inline-block; width:100%; border-bottom:1px solid #aaa; height:12px;"></span></td></tr>`;
            // Sezione note tradizionali
            html += `<tr class="logistic-row"><td colspan="14"><strong>Note:</strong><br><span style="display:inline-block; width:100%; border-bottom:1px solid #aaa; height:12px;"></span><br><span style="display:inline-block; width:100%; border-bottom:1px solid #aaa; height:12px;"></span><br><span style="display:inline-block; width:100%; border-bottom:1px solid #aaa; height:12px;"></span></td></tr>`;
            html += `</tbody></table>`;
            if (idx < selected.length - 1) {
                html += `<div style="page-break-after: always;"></div>`;
            }
        });
        html += `</body></html>`;
        const w = window.open('', '_blank');
        w.document.write(html);
        w.document.close();
        w.focus();
        setTimeout(() => { w.print(); }, 500);
        if (modal) modal.style.display = 'none';
    };
    document.addEventListener('DOMContentLoaded', () => {
        const packingBtn = document.getElementById('packingListBtn');
        if (packingBtn) {
            packingBtn.addEventListener('click', openPackingListModal);
        }
        const closeBtn = document.getElementById('packingListCloseBtn');
        if (closeBtn) {
            closeBtn.addEventListener('click', closePackingListModal);
        }
        const createBtn = document.getElementById('packingListCreateBtn');
        if (createBtn) {
            createBtn.addEventListener('click', buildPackingList);
        }
    });
    </script>

    <!-- ====================================================================
         AGGIUNTE DI FUNZIONALITÀ
         Questo blocco riporta le funzionalità introdotte nelle versioni
         successive (persistenza dello stato per la tabella Medical Device,
         tooltip per scatole/scatoloni, uniformità dello scroll, ordinamento
         multi‑colonna e trascinabilità dei modali) mantenendo invariata la
         logica della Packing List.  Le funzioni qui sotto non alterano il
         comportamento della Packing List originaria ma arricchiscono le altre
         sezioni dell’applicazione.
    ===================================================================== -->
    <script>
    (function() {
        /*
         * 1. Persistenza e rendering della tabella Medical Device
         *
         * Mantiene un unico stato (medicalDeviceState) in memoria e in
         * localStorage. Ogni riga ha un rowId stabile. L'aggiunta o la
         * modifica di una riga aggiorna lo stato; il rendering rigenera
         * completamente il tbody in base ai filtri correnti senza alterare lo
         * state. Le righe manuali vengono preservate tra le sessioni.
         */
        // Inizializza uno stato vuoto. Il caricamento dal localStorage viene
        // eseguito esplicitamente da loadMedicalDeviceState() dopo il login.
        window.medicalDeviceState = [];
        // Generatore di rowId univoci; utilizza timestamp e random.
        function generateRowId() {
            return 'md_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2,5);
        }
        /**
         * Carica la medicalDeviceState dal localStorage. Questo metodo viene
         * invocato durante l'inizializzazione post‑login per evitare
         * rallentamenti all'apertura della pagina. Dopo aver caricato i dati,
         * assegna un rowId a ciascun elemento se manca.
         */
        function loadMedicalDeviceState() {
            window.medicalDeviceState = [];
            try {
                const mdSaved = localStorage.getItem('medicalDeviceState');
                if (mdSaved) {
                    window.medicalDeviceState = JSON.parse(mdSaved);
                }
            } catch (e) {
                console.warn('Errore nel parsing di medicalDeviceState:', e);
                window.medicalDeviceState = [];
            }
            // Assegna un rowId a tutti gli elementi esistenti se manca
            const nowTs = Date.now().toString(36);
            window.medicalDeviceState.forEach(item => {
                if (!item.rowId) item.rowId = 'md_' + nowTs + '_' + Math.random().toString(36).substr(2,5);
            });
        }

        // Rende disponibile loadMedicalDeviceState a livello globale affinché
        // possa essere invocata dall'inizializzazione post‑login.  Senza questa
        // assegnazione, la funzione rimarrebbe confinata nello scope locale.
        window.loadMedicalDeviceState = loadMedicalDeviceState;
        // Salvataggio persistente dello stato
        function saveMedicalDeviceState() {
            try {
                localStorage.setItem('medicalDeviceState', JSON.stringify(medicalDeviceState));
            } catch (e) {
                console.warn('Impossibile salvare medicalDeviceState:', e);
            }
        }
        // Rendering della tabella filtrata
        function renderMedicalDeviceTable() {
            const body = window.medicalDeviceTableBody;
            if (!body) return;
            body.innerHTML = '';
            // Acquisisci filtri
            const startDate = (window.medicalDeviceStartDateInput && window.medicalDeviceStartDateInput._flatpickr) ? window.medicalDeviceStartDateInput._flatpickr.selectedDates[0] : null;
            const endDateOriginal = (window.medicalDeviceEndDateInput && window.medicalDeviceEndDateInput._flatpickr) ? window.medicalDeviceEndDateInput._flatpickr.selectedDates[0] : null;
            let endDate = endDateOriginal ? new Date(endDateOriginal.getTime()) : null;
            if (endDate) endDate.setHours(23,59,59,999);
            const filterCodiceText = (window.filterMedicalDeviceCodice && window.filterMedicalDeviceCodice.value || '').toLowerCase().trim();
            const filterDescrizioneText = (window.filterMedicalDeviceDescrizione && window.filterMedicalDeviceDescrizione.value || '').toLowerCase().trim();
            const filterClienteText = (window.filterMedicalDeviceCliente && window.filterMedicalDeviceCliente.value || '').toLowerCase().trim();
            // Recupera i nuovi filtri per data e lotto
            const filterDataText = (window.filterMedicalDeviceData && window.filterMedicalDeviceData.value || '').toLowerCase().trim();
            const filterLottoText = (window.filterMedicalDeviceLotto && window.filterMedicalDeviceLotto.value || '').toLowerCase().trim();
            // Filtra lo state
            const filtered = medicalDeviceState.filter(item => {
                // Filtro per data (produzioneData o data)
                let dateStr = item.produzioneData || item.data || '';
                if (startDate || endDate) {
                    const parts = String(dateStr).split('/');
                    if (parts.length === 3) {
                        const rowDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
                        if (startDate && rowDate < startDate) return false;
                        if (endDate && rowDate > endDate) return false;
                    }
                }
                // Filtro per codice/descrizione/cliente
                if (filterCodiceText && String(item.codice || '').toLowerCase().indexOf(filterCodiceText) === -1) return false;
                const descr = item.prodotto || item.descrizione || '';
                if (filterDescrizioneText && String(descr).toLowerCase().indexOf(filterDescrizioneText) === -1) return false;
                if (filterClienteText && String(item.cliente || '').toLowerCase().indexOf(filterClienteText) === -1) return false;
                // Filtro live per data: confronta il campo data o produzioneData con la stringa
                if (filterDataText) {
                    const dateCandidate = String(item.produzioneData || item.data || '').toLowerCase();
                    if (dateCandidate.indexOf(filterDataText) === -1) return false;
                }
                // Filtro live per lotto: confronta il lotto
                if (filterLottoText) {
                    if (String(item.lotto || '').toLowerCase().indexOf(filterLottoText) === -1) return false;
                }
                return true;
            });
            // Ricostruzione righe
            filtered.forEach(item => {
                const rowEl = window.createMedicalDeviceRow ? window.createMedicalDeviceRow(item, !!item.isManual) : null;
                if (!rowEl) return;
                rowEl.dataset.rowId = item.rowId;
                // Aggiorna lo stato quando gli input cambiano
                Array.from(rowEl.querySelectorAll('input')).forEach((inputEl, idx) => {
                    inputEl.addEventListener('change', () => {
                        const stateItem = medicalDeviceState.find(r => r.rowId === item.rowId);
                        if (!stateItem) return;
                        switch(idx) {
                            case 0: stateItem.codice = inputEl.value; break;
                            case 1: stateItem.prodotto = inputEl.value; break;
                            case 2: stateItem.cliente = inputEl.value; break;
                            case 3: stateItem.confezionamentoPezzi = inputEl.value; break;
                            case 4: stateItem.scarti = inputEl.value; break;
                            case 6: stateItem.siringhePerScatola = inputEl.value; break;
                            default: break;
                        }
                        saveMedicalDeviceState();
                    });
                });
                body.appendChild(rowEl);
            });
        }
        // Override della funzione originale per aggiornare la tabella
        if (typeof window.updateMedicalDeviceProductionTable === 'function') {
            const originalFn = window.updateMedicalDeviceProductionTable;
            window.updateMedicalDeviceProductionTable = function() {
                renderMedicalDeviceTable();
                // Non richiamiamo originalFn per evitare duplicazioni
            };
        }
        // Aggiunta nuova riga manuale: aggiorna stato e re-render
        if (window.addMedicalDeviceRowBtn) {
            window.addMedicalDeviceRowBtn.addEventListener('click', () => {
                const newItem = { rowId: generateRowId(), isManual: true, codice: '', prodotto: '', cliente: '', confezionamentoPezzi: '', scarti: '', quantitaDaProdurre: '', siringhePerScatola: '' };
                medicalDeviceState.push(newItem);
                saveMedicalDeviceState();
                renderMedicalDeviceTable();
            });
        }
        // Il rendering iniziale della tabella Medical Device viene
        // eseguito esplicitamente dopo il login tramite initializeAfterLogin(),
        // evitando così di rallentare il caricamento della pagina.
        /*
         * 2. Tooltip “scatole/scatoloni”
         *
         * Al passaggio del mouse sul codice o sulla descrizione nella tabella MD,
         * viene mostrato un tooltip che riporta il numero di scatole piene,
         * eventuale avanzo e, se definito, gli scatoloni e le scatole restanti.
         * I dati vengono recuperati dal deviceRefData salvato in localStorage.
         */
        (function() {
            const tooltip = document.createElement('div');
            tooltip.className = 'md-tooltip';
            tooltip.style.position = 'absolute';
            tooltip.style.pointerEvents = 'none';
            tooltip.style.background = '#fff';
            tooltip.style.border = '1px solid #ccc';
            tooltip.style.padding = '6px';
            tooltip.style.fontSize = '12px';
            tooltip.style.boxShadow = '0 2px 8px rgba(0,0,0,0.2)';
            tooltip.style.display = 'none';
            document.body.appendChild(tooltip);
            function getRefInfo(codice) {
                try {
                    const refs = JSON.parse(localStorage.getItem('deviceRefData') || '[]');
                    const key = String(codice || '').trim().toUpperCase();
                    return refs.find(ref => String(ref.codice || '').trim().toUpperCase() === key) || null;
                } catch (e) {
                    return null;
                }
            }
            function computeBoxes(totalPieces, piecesPerBox, boxesPerCarton) {
                const tot = parseFloat(totalPieces);
                const perBox = parseFloat(piecesPerBox);
                if (!tot || !perBox) return null;
                const fullBoxes = Math.floor(tot / perBox);
                const remainder = tot % perBox;
                if (boxesPerCarton) {
                    const perCarton = parseFloat(boxesPerCarton);
                    const cartons = Math.floor(fullBoxes / perCarton);
                    const remainingBoxes = fullBoxes % perCarton;
                    return { fullBoxes, remainder, cartons, remainingBoxes };
                }
                return { fullBoxes, remainder };
            }
            document.addEventListener('mouseover', (ev) => {
                const cell = ev.target.closest('#medicalDeviceProductionTable td');
                if (!cell) return;
                const rowEl = cell.parentElement;
                if (!rowEl) return;
                // Solo codice (col0) o descrizione (col1)
                const colIndex = Array.from(rowEl.children).indexOf(cell);
                if (colIndex !== 0 && colIndex !== 1) return;
                const rowId = rowEl.dataset.rowId;
                let item;
                if (rowId) {
                    item = medicalDeviceState.find(x => x.rowId === rowId);
                }
                if (!item) {
                    // Fallback: costruisci da input
                    const inputs = rowEl.querySelectorAll('input');
                    item = {
                        codice: inputs[0] ? inputs[0].value : '',
                        prodotto: inputs[1] ? inputs[1].value : '',
                        cliente: inputs[2] ? inputs[2].value : '',
                        scarti: inputs[4] ? inputs[4].value : '',
                        quantitaDaProdurre: '',
                        siringhePerScatola: inputs[6] ? inputs[6].value : ''
                    };
                }
                const refInfo = getRefInfo(item.codice);
                const piecesValue = item.quantitaDaProdurre || item.scarti || '';
                const perBox = refInfo && (refInfo.siringhePerScatola || refInfo.pezziPerScatola);
                const perCarton = refInfo && refInfo.scatolePerScatolone;
                const comp = computeBoxes(piecesValue, perBox, perCarton);
                let html = '';
                html += `<strong>Codice:</strong> ${item.codice || ''}<br>`;
                html += `<strong>Descrizione:</strong> ${item.prodotto || ''}<br>`;
                html += `<strong>Cliente:</strong> ${item.cliente || ''}<br>`;
                html += `<strong>Pezzi reali:</strong> ${piecesValue || 'N/D'}<br>`;
                html += `<strong>Pezzi/scatola:</strong> ${perBox || 'N/D'}<br>`;
                if (comp) {
                    html += `<strong>Scatole piene:</strong> ${comp.fullBoxes}<br>`;
                    html += `<strong>Avanzo:</strong> ${comp.remainder}<br>`;
                    if (comp.cartons !== undefined) {
                        html += `<strong>Scatoloni:</strong> ${comp.cartons}<br>`;
                        html += `<strong>Scatole restanti:</strong> ${comp.remainingBoxes}<br>`;
                    }
                } else {
                    html += `<em>Dati insufficienti per calcolare le scatole</em>`;
                }
                tooltip.innerHTML = html;
                const rect = cell.getBoundingClientRect();
                tooltip.style.left = (rect.right + 5 + window.scrollX) + 'px';
                tooltip.style.top = (rect.top + window.scrollY) + 'px';
                tooltip.style.display = 'block';
            });
            document.addEventListener('mouseout', (ev) => {
                if (ev.target.closest && ev.target.closest('#medicalDeviceProductionTable td')) {
                    tooltip.style.display = 'none';
                }
            });
        })();
        /*
         * 3. Uniformità dello scroll e header sticky
         *
         * Applica la classe .table-wrapper con overflow automatico a tutti i
         * contenitori di tabella e rende sticky tutte le intestazioni thead th.
         */
        (function() {
            // Aggiungi la classe a tutti i wrapper tabelle
            document.querySelectorAll('.daily-production-table-wrapper, .table-container, .table-wrapper').forEach(wrapper => {
                wrapper.classList.add('table-wrapper');
                wrapper.style.overflow = 'auto';
            });
            // Rendi sticky le intestazioni
            const ths = document.querySelectorAll('table thead th');
            ths.forEach(th => {
                th.style.position = 'sticky';
                th.style.top = '0';
                th.style.background = th.style.background || '#f7f7f7';
                th.style.zIndex = th.style.zIndex || '5';
            });
        })();
        /*
         * 4. Ordinamento multi‑colonna con icone a doppia freccia
         *
         * Per ogni tabella presente nella pagina, aggiunge una piccola icona
         * (↕) e gestisce i click sulle intestazioni per cicli di ordinamento
         * ascendente, discendente e nessuno. Con Shift+clic è possibile
         * applicare ordinamenti multipli su più colonne. Lo stato viene
         * salvato in localStorage per essere ripristinato al caricamento.
         */
        (function() {
            document.querySelectorAll('table').forEach((tbl, tableIndex) => {
                const headers = tbl.querySelectorAll('thead th');
                const sortStateKey = 'sortState_' + (tbl.id || tableIndex);
                let sortState = [];
                try {
                    const saved = localStorage.getItem(sortStateKey);
                    if (saved) sortState = JSON.parse(saved);
                } catch (e) {}
                headers.forEach((th, colIndex) => {
                    // Salta colonne non ordinabili
                    if (th.dataset.sortable === 'false') return;
                    // Aggiungi icona
                    const icon = document.createElement('span');
                    icon.className = 'sort-icon';
                    icon.style.marginLeft = '4px';
                    icon.textContent = '↕';
                    th.appendChild(icon);
                    th.style.cursor = 'pointer';
                    th.setAttribute('aria-sort', 'none');
                    // Funzione per aggiornare le icone
                    function updateIcons() {
                        headers.forEach((h, idx) => {
                            const rule = sortState.find(r => r.key === idx);
                            const iconEl = h.querySelector('.sort-icon');
                            if (!iconEl) return;
                            if (rule) {
                                if (rule.dir === 'asc') {
                                    h.setAttribute('aria-sort', 'ascending');
                                    iconEl.textContent = '↑';
                                } else {
                                    h.setAttribute('aria-sort', 'descending');
                                    iconEl.textContent = '↓';
                                }
                            } else {
                                h.setAttribute('aria-sort', 'none');
                                iconEl.textContent = '↕';
                            }
                        });
                    }
                    // Handler click
                    th.addEventListener('click', (ev) => {
                        const isShift = ev.shiftKey;
                        const existing = sortState.find(r => r.key === colIndex);
                        if (!existing) {
                            const newRule = { key: colIndex, dir: 'asc' };
                            if (isShift) {
                                sortState.push(newRule);
                            } else {
                                sortState = [newRule];
                            }
                        } else {
                            if (existing.dir === 'asc') {
                                existing.dir = 'desc';
                            } else if (existing.dir === 'desc') {
                                // Rimuove la regola
                                sortState = sortState.filter(r => r.key !== colIndex);
                            }
                            if (!isShift) {
                                // Porta in testa la colonna cliccata
                                if (existing.dir) {
                                    sortState = [existing];
                                }
                            }
                        }
                        updateIcons();
                        // Sorting righe
                        const tbody = tbl.querySelector('tbody');
                        if (!tbody) return;
                        const rows = Array.from(tbody.querySelectorAll('tr'));
                        const collator = new Intl.Collator('it', { numeric: true, sensitivity: 'base' });
                        rows.sort((a, b) => {
                            for (const rule of sortState) {
                                const ax = a.children[rule.key];
                                const bx = b.children[rule.key];
                                const aval = ax ? (ax.innerText || ax.querySelector('input')?.value || '') : '';
                                const bval = bx ? (bx.innerText || bx.querySelector('input')?.value || '') : '';
                                let comp = 0;
                                // Determina il tipo dal data-type o fallback testo
                                const type = ax ? (ax.dataset.type || 'text') : 'text';
                                if (type === 'number') {
                                    const aNum = parseFloat(String(aval).replace(/\./g,'').replace(',','.')) || 0;
                                    const bNum = parseFloat(String(bval).replace(/\./g,'').replace(',','.')) || 0;
                                    comp = aNum - bNum;
                                } else if (type === 'date') {
                                    const parseDate = (str) => {
                                        const p = String(str).split('/');
                                        if (p.length === 3) return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0])).getTime();
                                        const d = Date.parse(str);
                                        return isNaN(d) ? 0 : d;
                                    };
                                    comp = parseDate(aval) - parseDate(bval);
                                } else if (type === 'boolean') {
                                    const aBool = /^(true|si|sì|yes|1)$/i.test(String(aval)) ? 1 : 0;
                                    const bBool = /^(true|si|sì|yes|1)$/i.test(String(bval)) ? 1 : 0;
                                    comp = aBool - bBool;
                                } else {
                                    comp = collator.compare(String(aval), String(bval));
                                }
                                if (comp !== 0) {
                                    return rule.dir === 'asc' ? comp : -comp;
                                }
                            }
                            return 0;
                        });
                        rows.forEach(r => tbody.appendChild(r));
                        // Persisti lo stato
                        localStorage.setItem(sortStateKey, JSON.stringify(sortState));
                    });
                    // Aggiorna icona iniziale se la colonna è già presente nello stato
                    updateIcons();
                });
            });
        })();
        /*
         * 5. Trascinabilità dei modali e alert
         *
         * Rende trascinabili i principali modali/alert del sistema. Gli elementi
         * vengono spostati tramite drag&drop della loro intestazione (se
         * presente) o del contenuto stesso. Questo permette all'utente di
         * liberare la visuale dei dati sottostanti durante l'uso dell'app.
         */
        (function() {
            function makeDraggable(modal, handle) {
                let dragging = false;
                let startX = 0;
                let startY = 0;
                let origX = 0;
                let origY = 0;
                function onMove(e) {
                    if (!dragging) return;
                    const dx = e.clientX - startX;
                    const dy = e.clientY - startY;
                    modal.style.position = 'absolute';
                    modal.style.left = (origX + dx) + 'px';
                    modal.style.top = (origY + dy) + 'px';
                }
                function onUp() {
                    dragging = false;
                    document.removeEventListener('mousemove', onMove);
                    document.removeEventListener('mouseup', onUp);
                }
                (handle || modal).addEventListener('mousedown', (e) => {
                    if (e.button !== 0) return;
                    dragging = true;
                    startX = e.clientX;
                    startY = e.clientY;
                    const rect = modal.getBoundingClientRect();
                    origX = rect.left + window.scrollX;
                    origY = rect.top + window.scrollY;
                    document.addEventListener('mousemove', onMove);
                    document.addEventListener('mouseup', onUp);
                    e.preventDefault();
                });
            }
            // Applica la trascinabilità a elementi noti. Puoi aggiungere qui
            // altri selettori per modali personalizzati.
            document.addEventListener('DOMContentLoaded', () => {
                document.querySelectorAll('#adrNotification, .modal-dialog, .alert-modal, .tour-step').forEach(modal => {
                    const handle = modal.querySelector('.modal-header') || modal.querySelector('.adr-alert-content') || modal;
                    makeDraggable(modal, handle);
                });
            });
        })();
    })();
    </script>
    <!-- Script per la navigazione rapida: gestisce il click sui link del menu di
         navigazione e scrolla dolcemente alle sezioni corrispondenti.  Si
         esegue al DOMContentLoaded per assicurarsi che gli elementi target
         siano presenti nel DOM. -->
    <script>
    document.addEventListener('DOMContentLoaded', function() {
        var links = document.querySelectorAll('#quickNav a');
        links.forEach(function(link) {
            link.addEventListener('click', function(e) {
                e.preventDefault();
                var targetId = this.getAttribute('data-scroll-target');
                var section = document.getElementById(targetId);
                if (section) {
                    section.scrollIntoView({ behavior: 'smooth' });
                }
            });
        });
    });
    </script>
<script>
// === Script di supporto per la navigazione rapida e il registro Sblocchi CQ/QA ===
document.addEventListener('DOMContentLoaded', () => {
  // Navigazione verticale: scrolla dolcemente alla sezione indicata
  const quickNavVert = document.getElementById('quickNavVertical');
  if (quickNavVert) {
    quickNavVert.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        const targetId = link.getAttribute('data-scroll-target');
        const targetEl = document.getElementById(targetId);
        if (targetEl) {
          targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      });
    });
  }

  // Gestione del click sul pulsante di logout.  Rimuove il livello utente
  // dalla sessione e ricarica la pagina per mostrare nuovamente il pannello di login.
  const logoutNavBtn = document.getElementById('logoutNavBtn');
  if (logoutNavBtn) {
    logoutNavBtn.addEventListener('click', (ev) => {
      ev.preventDefault();
      try {
        sessionStorage.removeItem('userLevel');
      } catch (e) {
        console.warn('Impossibile cancellare userLevel dalla sessione:', e);
      }
      // Facoltativamente pulisci altre chiavi correlate se necessario
      location.reload();
    });
  }

  // Registro Sblocchi CQ/QA
  const sbloccoBtn = document.getElementById('sbloccoBtn');
  const sbloccoModal = document.getElementById('sbloccoModal');
  const sbloccoCloseBtn = document.getElementById('sbloccoCloseBtn');
  const sbloccoExportBtn = document.getElementById('sbloccoExportBtn');
  const sbloccoPrintBtn = document.getElementById('sbloccoPrintBtn');
  // Per il modulo legacy utilizziamo prefissi "legacy" per evitare conflitti.  Se gli elementi
  // legacy non esistono, queste variabili saranno null e le funzioni legacy non verranno eseguite.
  const sbloccoStartDateEl = document.getElementById('legacySbloccoStartDate');
  const sbloccoEndDateEl = document.getElementById('legacySbloccoEndDate');
  const sbloccoStateFilterEl = document.getElementById('legacySbloccoStateFilter');
  const sbloccoSearchInputEl = document.getElementById('legacySbloccoSearchInput');
  let sbloccoEventsData = [];

  function loadSbloccoEvents() {
    try {
      const saved = localStorage.getItem('sbloccoEventsData');
      sbloccoEventsData = saved ? JSON.parse(saved) : [];
    } catch (e) {
      sbloccoEventsData = [];
    }
  }
  function saveSbloccoEvents() {
    localStorage.setItem('sbloccoEventsData', JSON.stringify(sbloccoEventsData));
  }
  // Espone globalmente la funzione di registrazione di uno sblocco in modo che altri moduli possano richiamarla.
window.recordSbloccoEvent = function(eventObj) {
    // Salva gli eventi di sblocco utilizzando la nuova logica CQ/QA.  Se viene
    // specificato eventObj.type (o eventObj.entity) con valore 'QA' allora
    // l'evento viene memorizzato nella tabella QA, altrimenti di default nella tabella CQ.
    // Ogni evento viene arricchito con timestamp, data e ora per la visualizzazione.
    try {
      if (!window.sbloccoLoadData || !window.sbloccoSaveData) {
        console.warn('Funzioni di salvataggio sblocchi non disponibili');
        return;
      }
      const now = new Date();
      const typeRaw = (eventObj && (eventObj.type || eventObj.entity)) || '';
      const type = String(typeRaw).toUpperCase() === 'QA' ? 'QA' : 'CQ';
      const entry = {
        timestamp: now.toISOString(),
        date: now.toISOString().split('T')[0],
        dateTime: now.toLocaleString('it-IT'),
        ov: eventObj && eventObj.ov ? eventObj.ov : '',
        op: eventObj && eventObj.op ? eventObj.op : '',
        codice: eventObj && eventObj.codice ? eventObj.codice : '',
        descrizione: eventObj && eventObj.descrizione ? eventObj.descrizione : '',
        lotto: eventObj && eventObj.lotto ? eventObj.lotto : '',
        quantita: eventObj && (eventObj.quantita || eventObj.quantity) ? (eventObj.quantita || eventObj.quantity) : '',
        state: eventObj && (eventObj.state || eventObj.status) ? (eventObj.state || eventObj.status) : ''
      };
      const data = window.sbloccoLoadData(type) || [];
      data.push(entry);
      window.sbloccoSaveData(type, data);
    } catch (e) {
      console.warn('Errore nel salvataggio dell\'evento sblocco:', e);
    }
  };
  function renderSbloccoTable() {
    loadSbloccoEvents();
    applySbloccoFilters();
  }
  function applySbloccoFilters() {
    const startVal = sbloccoStartDateEl.value;
    const endVal = sbloccoEndDateEl.value;
    const stateVal = sbloccoStateFilterEl.value;
    const searchVal = sbloccoSearchInputEl.value.trim().toLowerCase();
    const tbody = document.querySelector('#legacySbloccoTable tbody');
    tbody.innerHTML = '';
    let total = 0, greenCount = 0, redCount = 0;
    sbloccoEventsData.forEach(ev => {
      let match = true;
      if (startVal && ev.date < startVal) match = false;
      if (endVal && ev.date > endVal) match = false;
      if (stateVal !== 'all' && ev.state !== stateVal) match = false;
      const searchable = [ev.ov, ev.op, ev.codice, ev.descrizione, ev.lotto, ev.quantita, ev.um].join(' ').toLowerCase();
      if (searchVal && !searchable.includes(searchVal)) match = false;
      if (match) {
        total++;
        if (ev.state === 'green') greenCount++;
        else if (ev.state === 'red') redCount++;
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${ev.dateTime || ''}</td>
          <td>${ev.ov || ''}</td>
          <td>${ev.op || ''}</td>
          <td>${ev.codice || ''}</td>
          <td>${ev.descrizione || ''}</td>
          <td>${ev.lotto || ''}</td>
          <td>${ev.quantita || ''} ${ev.um || ''}</td>
          <td>${ev.state || ''}</td>
        `;
        tbody.appendChild(tr);
      }
    });
    document.getElementById('legacySbloccoTotalCount').textContent = `Totale sblocchi: ${total}`;
    document.getElementById('legacySbloccoGreenCount').textContent = `Conformi: ${greenCount}`;
    document.getElementById('legacySbloccoRedCount').textContent = `Non conformi: ${redCount}`;
  }

  if (sbloccoBtn && sbloccoModal) {
    // Utilizza la nuova funzione di apertura del modale che provvede a
    // renderizzare le tabelle CQ/QA se disponibili.  In caso di assenza
    // della nuova funzione, la funzione openSbloccoModal effettuerà un
    // fallback sulla logica legacy.
    sbloccoBtn.addEventListener('click', () => {
      try {
        if (typeof openSbloccoModal === 'function') {
          openSbloccoModal();
        } else {
          // Fallback legacy: mostra semplicemente il modale e renderizza
          // tramite la vecchia funzione.
          sbloccoModal.style.display = 'flex';
          renderSbloccoTable();
        }
      } catch (e) {
        console.warn('Errore apertura modale sblocchi:', e);
      }
    });
  }
  if (sbloccoCloseBtn) {
    sbloccoCloseBtn.addEventListener('click', () => {
      // Utilizza la funzione unificata di chiusura del modale (se disponibile)
      try {
        if (typeof closeSbloccoModal === 'function') {
          closeSbloccoModal();
        } else if (sbloccoModal) {
          sbloccoModal.style.display = 'none';
        }
      } catch (e) {
        console.warn('Errore chiusura modale sblocchi:', e);
      }
    });
  }
  if (sbloccoStateFilterEl) sbloccoStateFilterEl.addEventListener('change', applySbloccoFilters);
  if (sbloccoStartDateEl) sbloccoStartDateEl.addEventListener('change', applySbloccoFilters);
  if (sbloccoEndDateEl) sbloccoEndDateEl.addEventListener('change', applySbloccoFilters);
  if (sbloccoSearchInputEl) sbloccoSearchInputEl.addEventListener('input', applySbloccoFilters);
  if (sbloccoExportBtn) {
    sbloccoExportBtn.addEventListener('click', () => {
      // Utilizza la funzione unificata di esportazione degli sblocchi.  Se
      // disponibile, exportSbloccoCSV si occuperà di gestire sia la versione
      // legacy che quella con doppia tabella CQ/QA.
      try {
        if (typeof exportSbloccoCSV === 'function') {
          exportSbloccoCSV();
        }
      } catch (e) {
        console.warn('Errore nella funzione di esportazione sblocchi:', e);
      }
    });
  }
  if (sbloccoPrintBtn) {
    sbloccoPrintBtn.addEventListener('click', () => {
      // Utilizza la funzione unificata di stampa.  printSbloccoTable
      // effettuerà il fallback alla logica legacy se necessario.
      try {
        if (typeof printSbloccoTable === 'function') {
          printSbloccoTable();
        }
      } catch (e) {
        console.warn('Errore nella funzione di stampa sblocchi:', e);
      }
    });
  }
});
</script>
    <!-- Tachimetro delle prestazioni: misura approssimativa del carico dati nel localStorage. -->
    <script>
    // Aggiorna il tachimetro delle prestazioni leggendo la dimensione dei dati
    // salvati nel localStorage.  Più dati vengono memorizzati (es. tabelle di
    // produzione, sblocchi, import), maggiore sarà il carico indicato.  Il
    // tachimetro è semplicemente una barra colorata che passa dal verde al
    // giallo al rosso e un'etichetta che mostra il peso totale in KB.  La
    // funzione viene eseguita inizialmente al caricamento della pagina e
    // successivamente ogni 10 secondi.
    /**
     * Aggiorna il tachimetro delle prestazioni.  Calcola la quantità di
     * caratteri salvati nel localStorage (che include i dati importati,
     * tabelle di produzione, sblocchi, ecc.) e ne deriva un rapporto
     * rispetto ad una soglia di 300k caratteri (~300 KB).  In base a tale
     * rapporto viene impostato il colore della barra: verde quando il
     * carico è basso, giallo quando si avvicina al limite di saturazione,
     * rosso quando si supera una soglia critica.  L'etichetta mostra
     * l'occupazione in KB con una sola cifra decimale.
     */
    function updatePerformanceGauge() {
        var size = 0;
        try {
            var dataStr = JSON.stringify(localStorage);
            size = dataStr ? dataStr.length : 0;
        } catch (e) {
            size = 0;
        }
        /*
         * Calcola il rapporto di utilizzo rispetto ad una soglia di 5 MB
         * (5 * 1024 * 1024 caratteri).  Questa soglia più elevata
         * riflette meglio il limite effettivo del localStorage (~5 MB per
         * dominio) e permette al tachimetro di muoversi gradualmente
         * anziché saturarsi immediatamente con file di grandi dimensioni.
         */
        var maxStorage = 5 * 1024 * 1024; // 5 MB
        var ratio = size / maxStorage;
        if (ratio > 1) ratio = 1;
        // Applica un ammorbidimento (smoothing) al valore mostrato per
        // evitare salti improvvisi dell'ago.  Manteniamo l'ultimo valore
        // calcolato in una variabile globale e lo aggiorniamo verso
        // l'obiettivo con un fattore di smorzamento (0.3).
        if (typeof window.lastGaugeRatio === 'undefined') {
            window.lastGaugeRatio = ratio;
        } else {
            window.lastGaugeRatio = window.lastGaugeRatio + (ratio - window.lastGaugeRatio) * 0.3;
        }
        var smoothed = window.lastGaugeRatio;
        var pointer = document.getElementById('gaugePointer');
        var label = document.getElementById('performanceGaugeLabel');
        if (!pointer || !label) return;
        // Calcola l'angolo di rotazione del puntatore: parte da -90° (0%) e
        // arriva a +90° (100%).  Clampa il rapporto a [0,1].
        var angle = -90 + smoothed * 180;
        pointer.style.transform = 'rotate(' + angle + 'deg)';
        label.textContent = 'Carico dati: ' + (size / 1024).toFixed(1) + ' KB';
    }

    // -----------------------------------------------------------------------------
    // Funzioni e strutture di supporto per il monitoraggio delle prestazioni
    // -----------------------------------------------------------------------------
    // Oggetto globale in cui vengono salvate le durate di rendering delle tabelle
    // (in millisecondi).  Le chiavi corrispondono ai tipi di tabella ('CQ' e 'QA').
    if (!window.lastRenderDurations) window.lastRenderDurations = {};

    /**
     * Aggiorna l'elemento della UI dedicato alle metriche di prestazione.
     * Viene chiamato dopo ogni renderTable() per visualizzare i tempi di
     * elaborazione.  Se non sono presenti metriche, il testo viene azzerato.
     */
    function updatePerfMetrics() {
        var el = document.getElementById('perfMetrics');
        if (!el) return;
        var parts = [];
        if (window.lastRenderDurations.CQ) parts.push('CQ ' + window.lastRenderDurations.CQ.toFixed(0) + ' ms');
        if (window.lastRenderDurations.QA) parts.push('QA ' + window.lastRenderDurations.QA.toFixed(0) + ' ms');
        el.textContent = parts.length ? ('Ultimo rendering: ' + parts.join(' | ')) : '';
    }
    // Programma l'aggiornamento periodico del tachimetro
    document.addEventListener('DOMContentLoaded', function() {
        // mostra il tachimetro solo se il contenitore esiste
        var perfWrapper = document.getElementById('performanceGaugeWrapper');
        if (perfWrapper) {
            perfWrapper.style.display = 'block';
        }
        updatePerformanceGauge();
        // Aggiorna il tachimetro ogni 15 secondi: frequenza più bassa per
        // ridurre l'overhead di lettura dal localStorage e migliorare la
        // reattività generale della pagina.
setInterval(updatePerformanceGauge, 15000);
    });
    </script>
</body>
<script>
// === Script per la navigazione verticale e il registro degli sblocchi CQ/QA ===
document.addEventListener('DOMContentLoaded', function() {
    // Gestione della navigazione verticale: scroll dolce su click dei numeri
    const vLinks = document.querySelectorAll('.quick-nav-vertical a');
    vLinks.forEach(function(link) {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const targetId = this.getAttribute('data-scroll-target');
            const section = document.getElementById(targetId);
            if (section) {
                section.scrollIntoView({ behavior: 'smooth' });
            }
        });
    });
    // Pulsante per aprire il registro degli sblocchi
    const sbloccoBtnEl = document.getElementById('sbloccoBtn');
    if (sbloccoBtnEl) {
        sbloccoBtnEl.addEventListener('click', function() {
            openSbloccoModal();
        });
    }
    // Pulsante per chiudere il registro degli sblocchi
    const sbloccoClose = document.getElementById('sbloccoCloseBtn');
    if (sbloccoClose) {
        sbloccoClose.addEventListener('click', function() {
            closeSbloccoModal();
        });
    }
    // Filtri per il registro degli sblocchi
    // Elementi legacy per il vecchio registro: usiamo id con prefisso "legacy".
    const startInput = document.getElementById('legacySbloccoStartDate');
    const endInput   = document.getElementById('legacySbloccoEndDate');
    const stateFilter = document.getElementById('legacySbloccoStateFilter');
    const searchInput = document.getElementById('legacySbloccoSearchInput');
    if (startInput) startInput.addEventListener('change', renderSbloccoTable);
    if (endInput) endInput.addEventListener('change', renderSbloccoTable);
    if (stateFilter) stateFilter.addEventListener('change', renderSbloccoTable);
    if (searchInput) searchInput.addEventListener('input', renderSbloccoTable);
    const exportBtn = document.getElementById('sbloccoExportBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', function() {
            exportSbloccoCSV();
        });
    }
    const printBtn = document.getElementById('sbloccoPrintBtn');
    if (printBtn) {
        printBtn.addEventListener('click', function() {
            printSbloccoTable();
        });
    }
});

// Registra un nuovo evento di sblocco CQ/QA.
function recordSbloccoEvent(event) {
    // Se sono disponibili le nuove funzioni di gestione degli sblocchi (CQ/QA),
    // delega la registrazione all'implementazione moderna.  Questo evita di
    // sovrascrivere il record con la logica legacy e permette di mantenere i
    // dati separati per CQ e QA.  In caso contrario, utilizza la vecchia
    // gestione basata su un singolo array di eventi.
    try {
        if (typeof window.sbloccoLoadData === 'function' && typeof window.sbloccoSaveData === 'function') {
            // Determina se l'evento specifica un tipo/entità; se non presente
            // assume CQ come default.  La funzione recordSbloccoEvent
            // moderna si occuperà di arricchire il dato con timestamp.
            const eventObj = event || {};
            // Se esiste una versione globale di recordSbloccoEvent diversa
            // da questa funzione locale, delega a quella per evitare
            // ricorsione infinita.  In caso contrario, prosegui con la logica legacy.
            if (typeof window.recordSbloccoEvent === 'function' && window.recordSbloccoEvent !== recordSbloccoEvent) {
                window.recordSbloccoEvent(eventObj);
                return;
            }
            // se la funzione globale coincide con questa locale, non tentare
            // di richiamarla di nuovo per evitare ricorsione.
            // Continua alla logica legacy qui sotto.
        }
    } catch (err) {
        console.warn('Errore durante il tentativo di salvare con la nuova logica sblocco:', err);
    }
    // Logica legacy: salva l'evento in un unico array in localStorage e,
    // eventualmente, invia i dati al server.  Questa sezione viene eseguita
    // solo se le nuove funzioni non sono disponibili.
    try {
        let list = [];
        try {
            list = JSON.parse(localStorage.getItem('sbloccoEventsData') || '[]');
        } catch (e) {
            list = [];
        }
        list.push(event);
        localStorage.setItem('sbloccoEventsData', JSON.stringify(list));
        if (typeof saveDataToServer === 'function') {
            saveDataToServer();
        }
    } catch (e) {
        console.warn('Errore nel salvataggio dell\'evento sblocco:', e);
    }
}

    // Rende accessibili globalmente le metriche di prestazione e la funzione di aggiornamento.
    if (typeof window !== 'undefined') {
        window.lastRenderDurations = window.lastRenderDurations || {};
        window.updatePerfMetrics = updatePerfMetrics;
    }

// Restituisce l'elenco completo degli eventi di sblocco
function getAllSbloccoEventsData() {
    try {
        const list = JSON.parse(localStorage.getItem('sbloccoEventsData') || '[]');
        return Array.isArray(list) ? list : [];
    } catch (e) {
        return [];
    }
}

function openSbloccoModal() {
    const modal = document.getElementById('sbloccoModal');
    if (modal) {
        modal.style.display = 'flex';
        // Se sono disponibili le nuove funzioni di rendering per le tabelle CQ e QA,
        // chiamale per popolare immediatamente le tabelle.  Questo garantisce che
        // eventuali modifiche allo stato vengano riflesse quando il modale viene
        // aperto.  Le chiamate sono protette da typeof per evitare errori in
        // ambienti legacy.
        if (typeof window.renderCQ === 'function') {
            try { window.renderCQ(); } catch(e) { console.warn('Errore rendering CQ:', e); }
        }
        if (typeof window.renderQA === 'function') {
            try { window.renderQA(); } catch(e) { console.warn('Errore rendering QA:', e); }
        }
        // Non chiamare il rendering legacy quando sono disponibili le nuove tabelle CQ/QA.
    }
}
function closeSbloccoModal() {
    const modal = document.getElementById('sbloccoModal');
    if (modal) modal.style.display = 'none';
}

// Funzione che applica i filtri e popola la tabella degli sblocchi
function renderSbloccoTable() {
    // Se esistono le nuove funzioni di rendering (CQ/QA) salta il rendering legacy
    if (typeof window.renderCQ === 'function' || typeof window.renderQA === 'function') {
        return;
    }
    const tbody = document.querySelector('#legacySbloccoTable tbody');
    if (!tbody) return;
    let events = getAllSbloccoEventsData();
    const stateVal = document.getElementById('legacySbloccoStateFilter')?.value || 'all';
    if (stateVal !== 'all') {
        events = events.filter(function(ev) { return ev.status === stateVal; });
    }
    const startVal = document.getElementById('legacySbloccoStartDate')?.value;
    const endVal   = document.getElementById('legacySbloccoEndDate')?.value;
    if (startVal) {
        const sd = new Date(startVal);
        events = events.filter(function(ev) {
            const d = new Date(ev.timestamp);
            return d >= sd;
        });
    }
    if (endVal) {
        const ed = new Date(endVal);
        ed.setHours(23,59,59,999);
        events = events.filter(function(ev) {
            const d = new Date(ev.timestamp);
            return d <= ed;
        });
    }
    const searchVal = document.getElementById('legacySbloccoSearchInput')?.value?.trim().toLowerCase() || '';
    if (searchVal) {
        events = events.filter(function(ev) {
            return (ev.ov || '').toLowerCase().includes(searchVal) ||
                   (ev.op || '').toLowerCase().includes(searchVal) ||
                   (ev.codice || '').toLowerCase().includes(searchVal) ||
                   (ev.descrizione || '').toLowerCase().includes(searchVal) ||
                   (ev.lotto || '').toLowerCase().includes(searchVal);
        });
    }
    var total = events.length;
    var greens = events.filter(function(ev) { return ev.status === 'green'; }).length;
    var reds = events.filter(function(ev) { return ev.status === 'red'; }).length;
    var totalSpan = document.getElementById('legacySbloccoTotalCount');
    var greenSpan = document.getElementById('legacySbloccoGreenCount');
    var redSpan   = document.getElementById('legacySbloccoRedCount');
    if (totalSpan) totalSpan.textContent = 'Totale sblocchi: ' + total;
    if (greenSpan) greenSpan.textContent = 'Conformi: ' + greens;
    if (redSpan) redSpan.textContent = 'Non conformi: ' + reds;
    events.sort(function(a,b) {
        return new Date(b.timestamp) - new Date(a.timestamp);
    });
    tbody.innerHTML = '';
    events.forEach(function(ev) {
        var tr = document.createElement('tr');
        var d = new Date(ev.timestamp);
        var dateStr = d.toLocaleString('it-IT');
        tr.innerHTML = '<td>' + dateStr + '</td>' +
                       '<td>' + (ev.ov || '') + '</td>' +
                       '<td>' + (ev.op || '') + '</td>' +
                       '<td>' + (ev.codice || '') + '</td>' +
                       '<td>' + (ev.descrizione || '') + '</td>' +
                       '<td>' + (ev.lotto || '') + '</td>' +
                       '<td>' + ((ev.quantita || '') + ' ' + (ev.um || '')).trim() + '</td>' +
                       '<td>' + ev.status + '</td>';
        tbody.appendChild(tr);
    });
}

// Esporta gli eventi filtrati in CSV
function exportSbloccoCSV() {
    // Se sono disponibili le nuove funzioni di esportazione (gestione CQ/QA), usa quelle
    if (typeof window.exportBoth === 'function') {
        window.exportBoth();
        return;
    }
    // Altrimenti esporta la tabella legacy (#sbloccoTable) come CSV singolo
    const tbody = document.querySelector('#legacySbloccoTable tbody');
    if (!tbody) return;
    const rows = Array.from(tbody.querySelectorAll('tr'));
    let csv = 'Data/Ora,OV,OP,Codice,Descrizione,Lotto,Quantità,Stato\n';
    rows.forEach(function(row) {
        const cols = Array.from(row.children).map(function(td) {
            var text = td.textContent || '';
            return '"' + text.replace(/"/g, '""') + '"';
        });
        csv += cols.join(',') + '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'sblocco_events.csv';
    a.click();
    URL.revokeObjectURL(url);
}

// Stampa la tabella degli sblocchi aprendo una nuova finestra
function printSbloccoTable() {
    // Se esistono le nuove funzioni di stampa per entrambe le tabelle, usa quelle
    if (typeof window.printTables === 'function') {
        window.printTables();
        return;
    }
    // Altrimenti stampa la tabella legacy (#sbloccoTable) in una nuova finestra
    const table = document.getElementById('legacySbloccoTable');
    if (!table) return;
    const newWin = window.open('', '_blank');
    const html = '<html><head><title>Registro Sblocchi</title>' +
                 '<style>body { font-family: Arial, sans-serif; margin: 20px; } table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #ccc; padding: 6px; font-size: 12px; white-space: nowrap; } th { background: #f5f5f5; }</style>' +
                 '</head><body>' +
                 '<h2>Registro Sblocchi CQ/QA</h2>' +
                 table.outerHTML +
                 '</body></html>';
    newWin.document.write(html);
    newWin.document.close();
    newWin.focus();
    setTimeout(function() { newWin.print(); }, 300);
}
</script>
</html>
<script>
// === PATCH: Warehouse Gantt dual-scroll + side buttons (top bar between Arrivi and Gantt) ===
(function(){
  if (window.__warehouseGanttScrollPatched__) return;
  window.__warehouseGanttScrollPatched__ = true;

  function ensureTopBarPlaced() {
    try {
      const topBar = document.getElementById('warehouseGanttExternalScrollbar');
      const ganttContainer = document.getElementById('warehouseGanttChartContainer');
      const arrivalsContainer = document.getElementById('arrivalScheduleContainer');
      if (!topBar || !ganttContainer) return;
      // Place topBar right before the Gantt container (i.e., between Arrivi and Gantt)
      const parent = ganttContainer.parentElement;
      if (parent && parent.children) {
        // If topBar is not immediately before ganttContainer, move it
        if (topBar.nextElementSibling !== ganttContainer) {
          parent.insertBefore(topBar, ganttContainer);
        }
      }
      // Make sure the top bar is visible
      topBar.style.display = 'block';
    } catch(e){ console.warn('ensureTopBarPlaced error:', e); }
  }

  function initWarehouseGanttScrollUX() {
    try {
      ensureTopBarPlaced();

      const wrapper = document.getElementById('warehouseGanttScrollWrapper');
      const topBar = document.getElementById('warehouseGanttExternalScrollbar');
      const sizer = topBar ? topBar.querySelector('.gantt-external-sizer') : null;
      const btnWrap = document.getElementById('warehouseGanttScrollButtonsWrapper');
      const btnLeft = document.getElementById('warehouseGanttScrollLeftBtn');
      const btnRight = document.getElementById('warehouseGanttScrollRightBtn');
      const container = document.getElementById('warehouseGanttChartContainer');

      if (!wrapper || !topBar || !sizer || !container || !btnLeft || !btnRight) return;

      // Size the sizer to Gantt total width
      function refreshSizer() {
        try {
          const totalW = Math.max(wrapper.scrollWidth || 0, container.clientWidth || 0);
          sizer.style.width = totalW + 'px';
        } catch(e){}
      }

      let lockA = false, lockB = false;
      function syncTopToBottom() {
        if (lockA) return;
        lockB = true;
        wrapper.scrollLeft = topBar.scrollLeft;
        updateButtons();
        lockB = false;
      }
      function syncBottomToTop() {
        if (lockB) return;
        lockA = true;
        topBar.scrollLeft = wrapper.scrollLeft;
        updateButtons();
        lockA = false;
      }

      function updateButtons() {
        const maxScroll = (wrapper.scrollWidth - wrapper.clientWidth) || 0;
        const x = wrapper.scrollLeft || 0;
        if (btnLeft) btnLeft.disabled = (x <= 0);
        if (btnRight) btnRight.disabled = (x >= maxScroll - 1);
      }

      function scrollStep(dir) {
        const step = Math.max(200, Math.floor(container.clientWidth * 0.8));
        const target = Math.max(0, Math.min(wrapper.scrollLeft + dir * step, (wrapper.scrollWidth - wrapper.clientWidth)));
        wrapper.scrollTo({ left: target, behavior: 'smooth' });
      }

      // Reposition side buttons vertically centered over the visible middle of the Gantt container
      function repositionButtons() {
        if (!btnWrap) return;
        const rect = container.getBoundingClientRect();
        // Hide when Gantt section is off-screen
        if (rect.bottom < 0 || rect.top > window.innerHeight) {
          btnWrap.style.display = 'none';
          return;
        }
        btnWrap.style.display = 'flex';
        const visibleTop = Math.max(rect.top, 0);
        const visibleBottom = Math.min(rect.bottom, window.innerHeight);
        const visibleHeight = visibleBottom - visibleTop;
        const btnH = btnWrap.offsetHeight || 0;
        const top = visibleTop + (visibleHeight - btnH) / 2;
        btnWrap.style.top = top + 'px';
      }

      // Wire events (avoid double-listeners via dataset flag)
      if (!topBar.dataset._patched) {
        topBar.addEventListener('scroll', syncTopToBottom, { passive: true });
        topBar.dataset._patched = '1';
      }
      if (!wrapper.dataset._patched) {
        wrapper.addEventListener('scroll', syncBottomToTop, { passive: true });
        wrapper.dataset._patched = '1';
      }
      if (!btnLeft.dataset._patched) {
        btnLeft.addEventListener('click', () => scrollStep(-1));
        btnLeft.dataset._patched = '1';
      }
      if (!btnRight.dataset._patched) {
        btnRight.addEventListener('click', () => scrollStep(1));
        btnRight.dataset._patched = '1';
      }
      window.addEventListener('resize', () => { refreshSizer(); repositionButtons(); updateButtons(); });
      window.addEventListener('scroll', () => { repositionButtons(); });

      // Initial layout
      refreshSizer();
      syncBottomToTop();
      repositionButtons();
      updateButtons();

      // Expose for re-use
      window.refreshWarehouseGanttScrollUX = () => {
        refreshSizer();
        syncBottomToTop();
        repositionButtons();
        updateButtons();
      };
    } catch(e){
      console.warn('initWarehouseGanttScrollUX error:', e);
    }
  }

  // Call on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      setTimeout(initWarehouseGanttScrollUX, 0);
    });
  } else {
    setTimeout(initWarehouseGanttScrollUX, 0);
  }

  /*
   * =====================================================================
   * Funzioni e inizializzazioni per la tabella "Merce in Scadenza".
   * Queste routine gestiscono la creazione delle righe, la lettura dei dati,
   * l'importazione dei file inventario, il filtraggio e l'esportazione.
   * Sono inserite qui alla fine dell'IIFE per garantire che tutte le
   * dipendenze (flatpickr, makeTableResizable, autoSaveAllData, ecc.)
   * siano già definite.  Gli eventi vengono agganciati al DOM una volta
   * caricato completamente.  La tabella con id "expiringGoodsTable" e
   * i relativi pulsanti sono definiti nel markup HTML.
   */
  // Riferimento al body della tabella "Merce in Scadenza"
  const expiringGoodsTableBody = document.querySelector('#expiringGoodsTable tbody');

  function createExpiringGoodsRow(rowData = {}) {
    if (!expiringGoodsTableBody) return null;
    const row = document.createElement('tr');
    const esc = (str) => String(str || '').replace(/"/g, '&quot;');
    row.innerHTML = `
            <td><input type="checkbox" class="expiring-row-selector"></td>
            <td><input type="text" value="${esc(rowData.codice || '')}"></td>
            <td><input type="text" value="${esc(rowData.articolo || '')}" style="text-align:left;"></td>
            <td><input type="text" value="${esc(rowData.lotto || '')}"></td>
            <td><input type="text" class="datepicker" value="${esc(rowData.scadenza || '')}"></td>
            <td><input type="number" value="${esc(rowData.quantita || '')}"></td>
            <td><input type="text" value="${esc(rowData.um || '')}"></td>
            <td><input type="text" value="${esc(rowData.layout || '')}"></td>
            <td><input type="text" value="${esc(rowData.famiglia || '')}" style="text-align:left;"></td>
            <td><input type="text" value="${esc(rowData.linea || '')}" style="text-align:left;"></td>
        `;
    row.querySelectorAll('.datepicker').forEach(input => {
      flatpickr(input, { dateFormat: 'd/m/Y', locale: 'it' });
    });
    row.querySelectorAll('input').forEach(input => {
      input.addEventListener('change', () => {
        autoSaveAllData();
      });
    });
    expiringGoodsTableBody.appendChild(row);
    return row;
  }

  function getExpiringGoodsRowData(row) {
    const cells = row.cells;
    return {
      codice: cells[1] && cells[1].querySelector('input') ? cells[1].querySelector('input').value : '',
      articolo: cells[2] && cells[2].querySelector('input') ? cells[2].querySelector('input').value : '',
      lotto: cells[3] && cells[3].querySelector('input') ? cells[3].querySelector('input').value : '',
      scadenza: cells[4] && cells[4].querySelector('input') ? cells[4].querySelector('input').value : '',
      quantita: cells[5] && cells[5].querySelector('input') ? cells[5].querySelector('input').value : '',
      um: cells[6] && cells[6].querySelector('input') ? cells[6].querySelector('input').value : '',
      layout: cells[7] && cells[7].querySelector('input') ? cells[7].querySelector('input').value : '',
      famiglia: cells[8] && cells[8].querySelector('input') ? cells[8].querySelector('input').value : '',
      linea: cells[9] && cells[9].querySelector('input') ? cells[9].querySelector('input').value : ''
    };
  }

  function getAllExpiringGoodsData() {
    const data = [];
    document.querySelectorAll('#expiringGoodsTable tbody tr').forEach(row => {
      data.push(getExpiringGoodsRowData(row));
    });
    return data;
  }

  function populateExpiringGoodsTable(data) {
    if (!expiringGoodsTableBody) return;
    expiringGoodsTableBody.innerHTML = '';
    data.forEach(rowData => {
      createExpiringGoodsRow(rowData);
    });
    const expTable = document.getElementById('expiringGoodsTable');
    if (expTable) {
      // Le funzioni per rendere ridimensionabile e ordinabile la tabella potrebbero non essere
      // definite in alcuni contesti (versioni del file base). Utilizziamo l'oggetto window per
      // verificare la loro presenza prima di invocarle, così da evitare errori di riferimento.
      if (typeof window !== 'undefined' && typeof window.makeTableResizable === 'function') {
        window.makeTableResizable(expTable);
      }
      if (typeof window !== 'undefined' && typeof window.makeTableSortable === 'function') {
        window.makeTableSortable(expTable);
      }
    }
  }

  function applyExpiringFilter() {
    const filterColSel = document.getElementById('filterExpiringColumn');
    const filterValInput = document.getElementById('filterExpiringValue');
    const startDateInput = document.getElementById('expiringStartDate');
    const endDateInput = document.getElementById('expiringEndDate');
    if (!filterColSel || !filterValInput || !startDateInput || !endDateInput) return;
    const filterCol = filterColSel.value;
    const filterVal = filterValInput.value.trim().toLowerCase();
    const startDate = flatpickr.parseDate(startDateInput.value, 'd/m/Y');
    const endDate = flatpickr.parseDate(endDateInput.value, 'd/m/Y');
    if (endDate) endDate.setHours(23, 59, 59, 999);
    document.querySelectorAll('#expiringGoodsTable tbody tr').forEach(row => {
      let showRow = true;
      const rowData = getExpiringGoodsRowData(row);
      if (filterCol && filterVal && !String(rowData[filterCol] || '').toLowerCase().includes(filterVal)) {
        showRow = false;
      }
      if (showRow && startDate) {
        const parts = String(rowData.scadenza || '').split('/');
        if (parts.length === 3) {
          const rDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          if (rDate < startDate || (endDate && rDate > endDate)) {
            showRow = false;
          }
        } else {
          showRow = false;
        }
      }
      row.style.display = showRow ? '' : 'none';
    });
  }

  async function processInventoryFile(file) {
    // Scrive un log dell'importazione solo se la funzione addLogEntry è disponibile
    // Utilizza l'oggetto global window per accedere a addLogEntry poiché in alcune build
    // il simbolo potrebbe essere definito solo a livello globale. In questo modo
    // evitiamo errori di riferimento non definito.
    if (typeof window !== 'undefined' && typeof window.addLogEntry === 'function') {
      window.addLogEntry(`--- Inizio importazione Inventario da: ${file.name} ---`);
    }
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
      const excelColToIndex = (col) => {
        let index = 0;
        for (let i = 0; i < col.length; i++) {
          index = index * 26 + col.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
        }
        return index - 1;
      };
      const colMap = {
        codice: excelColToIndex('C'),
        articolo: excelColToIndex('D'),
        layout: excelColToIndex('E'),
        lotto: excelColToIndex('F'),
        scadenza: excelColToIndex('G'),
        quantita: excelColToIndex('I'),
        um: excelColToIndex('J'),
        famiglia: excelColToIndex('Q'),
        linea: excelColToIndex('R')
      };
      // Espressione regolare aggiornata per intercettare sia forme singolari che plurali e
      // la variante con sigla (PMC). Esempi ammessi: "materie prime", "materia prima",
      // "principi attivi", "principio attivo", "dispositivo medico", "dispositivi medici",
      // "prodotti finiti cosmetici", "prodotto finito cosmetico", "integratori",
      // "integratore", "presidi medici chirurgici", "presidio medico chirurgico",
      // "presidi medici chirurgici (PMC)", "cosmetici (BULK)", "cosmetico (BULK)",
      // "miscellanea", "campioni laboratorio", "campione laboratorio". Le varianti
      // possono essere scritte in maiuscolo, minuscolo o misto.
      const validFamilyRegex = /(materi[ae]\s+prim[ae])|(princip[io]i?\s+attiv[io]i?)|(dispositiv[oi]\s+medic[io]i?)|(prodott[oi]\s+finit[oi]\s+cosmetic[oi])|(integrator[ei])|(presidi[o]?\s+medic[i]?\s+chirurgic[io](?:\s*\(PMC\))?)|(cosmetic[oi]\s*\(bulk\))|(miscellanea)|(campion[ei]\s+laboratorio)/i;
      const twoMonthsFromNow = new Date();
      twoMonthsFromNow.setMonth(twoMonthsFromNow.getMonth() + 2);
      twoMonthsFromNow.setHours(0, 0, 0, 0);
      const imported = [];
      // Variabile per conteggiare il numero di righe dati (non vuote) presenti nel file
      let totalRows = 0;
      jsonData.forEach((row, index) => {
        // Salta l'intestazione o le righe completamente vuote
        if (index === 0 || row.every(cell => cell === '')) return;
        // Conta la riga come valida ai fini del totale
        totalRows++;
        const fam = String(row[colMap.famiglia] || '').trim();
        if (!validFamilyRegex.test(fam)) return;

        const scadRawVal = row[colMap.scadenza];
        let scadDate;
        let scadDateStr = '';
        // Gestione delle date in formato seriale numerico di Excel o stringhe numeriche
        if (typeof scadRawVal === 'number' ||
            (typeof scadRawVal === 'string' && scadRawVal.trim() !== '' && !isNaN(scadRawVal) && !scadRawVal.includes('/') && !scadRawVal.includes('.'))) {
          const excelDays = parseFloat(scadRawVal);
          if (!isNaN(excelDays)) {
            // Epoch di Excel (sistema 1900): 1899-12-30 per compatibilità
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            scadDate = new Date(excelEpoch.getTime() + excelDays * 86400000);
            scadDate.setHours(0, 0, 0, 0);
            scadDateStr = scadDate.toLocaleDateString('it-IT');
          }
        } else {
          // Prova a interpretare la stringa utilizzando parseDateValue
          const scadStr = String(scadRawVal || '').trim();
          scadDateStr = parseDateValue(scadStr);
          const parts = scadDateStr.split('/');
          if (parts.length === 3) {
            scadDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          } else {
            scadDate = null;
          }
        }
        // Se la data non è stata interpretata correttamente salta la riga
        if (!scadDate) return;
        // Scarta le righe con data di scadenza successiva a due mesi da oggi
        if (scadDate > twoMonthsFromNow) return;
        const rowData = {
          codice: String(row[colMap.codice] || '').trim(),
          articolo: String(row[colMap.articolo] || '').trim(),
          lotto: String(row[colMap.lotto] || '').trim(),
          scadenza: scadDateStr,
          quantita: parseNumericValue(row[colMap.quantita]),
          um: String(row[colMap.um] || '').trim(),
          layout: String(row[colMap.layout] || '').trim(),
          famiglia: fam,
          linea: String(row[colMap.linea] || '').trim()
        };
        imported.push(rowData);
      });
      populateExpiringGoodsTable(imported);
      localStorage.setItem('expiring_goods_data', JSON.stringify(imported));
      if (typeof formatDateTimeForDisplay === 'function') {
        const nowStr = formatDateTimeForDisplay(new Date());
        localStorage.setItem('lastImportInventory', nowStr);
      } else {
        localStorage.setItem('lastImportInventory', Date.now().toString());
      }
      updateImportTimestamps();
      // Prova a salvare i dati sul server ma non bloccare l'importazione se fallisce
      try {
        if (typeof saveDataToServer === 'function') {
          await saveDataToServer();
        }
      } catch (e) {
        console.warn('saveDataToServer fallita durante import inventario:', e);
      }
      // Prova a salvare in locale tutti i dati; eventuali errori di quota vengono gestiti internamente
      try {
        autoSaveAllData();
      } catch (e) {
        console.warn('autoSaveAllData fallita durante import inventario:', e);
      }
      // Aggiorna il log e il messaggio di riepilogo includendo il totale delle righe non vuote, se disponibile
      if (typeof window !== 'undefined' && typeof window.addLogEntry === 'function') {
        window.addLogEntry(`Importazione Inventario completata. Righe totali: ${totalRows}. Righe importate: ${imported.length}.`);
      }
      await showAlert(`Importazione completata. Righe totali nel file: ${totalRows}, righe importate nella tabella Merce in Scadenza: ${imported.length}.`);
    } catch (error) {
      console.error('Errore durante l\'importazione inventario:', error);
      // Registra l'errore di importazione solo se addLogEntry è disponibile
      if (typeof window !== 'undefined' && typeof window.addLogEntry === 'function') {
        window.addLogEntry(`Importazione Inventario fallita: ${error.message}`);
      }
      await showAlert(`Errore durante l'importazione del file Inventario: ${error.message}.`);
    }
  }

  function loadExpiringGoodsDataFromLocal() {
    const local = localStorage.getItem('expiring_goods_data');
    if (local) {
      try {
        const data = JSON.parse(local);
        populateExpiringGoodsTable(data);
      } catch (e) {}
    }
  }

  // Inizializza filtri, pulsanti e carica dati quando il DOM è pronto
  document.addEventListener('DOMContentLoaded', () => {
    const startExp = document.getElementById('expiringStartDate');
    const endExp = document.getElementById('expiringEndDate');
    if (startExp && endExp) {
      flatpickr(startExp, { dateFormat: 'd/m/Y', locale: 'it', onChange: applyExpiringFilter });
      flatpickr(endExp, { dateFormat: 'd/m/Y', locale: 'it', onChange: applyExpiringFilter });
    }
    const filterColSel = document.getElementById('filterExpiringColumn');
    const filterValInp = document.getElementById('filterExpiringValue');
    const clearFilterBtn = document.getElementById('clearExpiringFilterBtn');
    if (filterColSel) filterColSel.addEventListener('change', applyExpiringFilter);
    if (filterValInp) filterValInp.addEventListener('input', applyExpiringFilter);
    if (clearFilterBtn) clearFilterBtn.addEventListener('click', () => {
      if (filterColSel) filterColSel.value = '';
      if (filterValInp) filterValInp.value = '';
      applyExpiringFilter();
    });
    const clearDateBtn = document.getElementById('clearExpiringDateBtn');
    if (clearDateBtn) clearDateBtn.addEventListener('click', () => {
      if (startExp && startExp._flatpickr) startExp._flatpickr.clear();
      if (endExp && endExp._flatpickr) endExp._flatpickr.clear();
      applyExpiringFilter();
    });
    const addBtn = document.getElementById('addExpiringRowBtn');
    const dupBtn = document.getElementById('duplicateExpiringRowBtn');
    const delBtn = document.getElementById('deleteExpiringRowBtn');
    if (addBtn) addBtn.addEventListener('click', () => {
      createExpiringGoodsRow({});
      autoSaveAllData();
    });
    if (dupBtn) dupBtn.addEventListener('click', () => {
      const selected = document.querySelectorAll('.expiring-row-selector:checked');
      if (selected.length === 0) {
        showAlert('Seleziona almeno una riga da duplicare nella tabella Merce in Scadenza.');
        return;
      }
      selected.forEach(chk => {
        const data = getExpiringGoodsRowData(chk.closest('tr'));
        createExpiringGoodsRow(data);
      });
      autoSaveAllData();
    });
    if (delBtn) delBtn.addEventListener('click', async () => {
      const selected = document.querySelectorAll('.expiring-row-selector:checked');
      if (selected.length === 0) {
        await showAlert('Seleziona almeno una riga da eliminare dalla tabella Merce in Scadenza.');
        return;
      }
      const confirmed = await showConfirm(`Sei sicuro di voler eliminare ${selected.length} riga/e selezionata/e dalla tabella Merce in Scadenza?`);
      if (confirmed) {
        selected.forEach(chk => {
          chk.closest('tr').remove();
        });
        autoSaveAllData();
        await showAlert('Riga/e eliminata/e con successo dalla tabella Merce in Scadenza.');
      }
    });
    const importBtn = document.getElementById('importInventoryBtn');
    const fileInput = document.getElementById('inventoryInput');
    if (importBtn && fileInput) {
      importBtn.addEventListener('click', () => {
        fileInput.click();
      });
      fileInput.addEventListener('change', async (event) => {
        const file = event.target.files && event.target.files[0];
        if (file) {
          await processInventoryFile(file);
          fileInput.value = '';
        }
      });
    }
    const exportBtn = document.getElementById('exportExpiringDataBtn');
    if (exportBtn) {
      exportBtn.addEventListener('click', () => {
        const data = getAllExpiringGoodsData();
        if (data.length === 0) return;
        const headers = Object.keys(data[0]);
        let csvContent = headers.join(';') + '\n';
        data.forEach(row => {
          csvContent += headers.map(h => `"${row[h]}"`).join(';') + '\n';
        });
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        const todayStr = new Date().toLocaleDateString('it-IT').replace(/\//g, '-');
        link.download = `merce_in_scadenza_${todayStr}.csv`;
        link.click();
      });
    }
    const emailBtn = document.getElementById('sendExpiringEmailBtn');
    if (emailBtn) {
      emailBtn.addEventListener('click', () => {
        showAlert("Funzionalità 'Invia Mail' per Merce in Scadenza da implementare.");
      });
    }
    loadExpiringGoodsDataFromLocal();
    const expTable = document.getElementById('expiringGoodsTable');
    if (expTable) {
      // Verifica che le funzioni makeTableResizable e makeTableSortable siano definite
      if (typeof window !== 'undefined' && typeof window.makeTableResizable === 'function') {
        window.makeTableResizable(expTable);
      }
      if (typeof window !== 'undefined' && typeof window.makeTableSortable === 'function') {
        window.makeTableSortable(expTable);
      }
    }
  });

  // Re-initialize after (re)draws of the Gantt if those hooks exist
  const _uwgc = window.updateWarehouseGanttChart;
  if (typeof _uwgc === 'function' && !_uwgc.__wrapped) {
    window.updateWarehouseGanttChart = function() {
      try { return _uwgc.apply(this, arguments); }
      finally {
        try { setTimeout(() => window.refreshWarehouseGanttScrollUX && window.refreshWarehouseGanttScrollUX(), 0); } catch(e){}
      }
    };
    window.updateWarehouseGanttChart.__wrapped = true;
  }
})();
// === END PATCH ===
// Ripristina automaticamente il login salvato in sessionStorage al caricamento della pagina.
(function() {
    try {
        const lvlStr = (typeof sessionStorage !== 'undefined') ? sessionStorage.getItem('userLevel') : null;
        const lvl = lvlStr ? parseInt(lvlStr, 10) : 0;
        if (lvl > 0) {
            // Se possibile, aggiorna la variabile globale currentUserLevel
            try {
                window.currentUserLevel = lvl;
            } catch (_) {}
            // Nascondi l'overlay di login, se presente
            const overlay = document.getElementById('loginOverlay');
            if (overlay) overlay.style.display = 'none';
            // Applica i permessi ed eventuali aggiornamenti della UI
            if (typeof applyPermissions === 'function') {
                try { applyPermissions(lvl); } catch (e) {}
            }
            // Aggiorna la tabella di analisi e inizializza dopo il login se le funzioni esistono.
            try { if (typeof updateAnalisiTable === 'function') updateAnalisiTable(); } catch (e) {}
            try { if (typeof initializeAfterLogin === 'function') initializeAfterLogin(); } catch (e) {}
        }
    } catch (e) {
        console.warn("Errore durante il ripristino dell'utente dalla sessionStorage:", e);
    }
})();
</script>
