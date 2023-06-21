#include "window.h"
#include <QPushButton>
#include <QApplication>
#include <QHBoxLayout>
#include <QString>
#include <QFileDialog>
#include <QLineEdit>
#include <QValidator>
#include <QLabel>
#include <QVector>
#include <QThread>
#include <QTableWidget>
#include <QtConcurrent/QtConcurrent>
#include <QFuture>
#include <QDebug>
#include <QHeaderView>
#include <QMessageBox>


#include <filesystem>
#include "xlsxdocument.h"
//#include "xlsxchartsheet.h"
//#include "xlsxcellrange.h"
//#include "xlsxchart.h"
//#include "xlsxrichstring.h"
//#include "xlsxworkbook.h"


Window::Window(QWidget *parent) : QWidget(parent) {
    // Set window parameters
    resize(350,200);
    setWindowTitle("Excel Parser");

    // Create import button
    import_button = new QPushButton("Import File");
    import_button->setObjectName("import_button");
    import_button->setCheckable(true);
    import_button->setFixedHeight(50);
    import_button->setMinimumWidth(350);
    import_button->setMaximumWidth(1000);

    // Create import progress label
    import_progress = new QLabel("Importing...");
    import_progress->hide();

    // Create row input field
    auto row_label = new QLabel("Row IDs:");
    row_input = new QLineEdit();
    row_input->setPlaceholderText("2,5,8");

    // Create table
    display = new QTableWidget();
    display->setVerticalScrollMode(QAbstractItemView::ScrollPerPixel);
    display->setHorizontalScrollMode(QAbstractItemView::ScrollPerPixel);

    //Creating layout
    auto main_layout = new QVBoxLayout(this);
    auto import_layout = new QVBoxLayout();
    auto row_input_layout = new QHBoxLayout();
    auto table_layout = new QHBoxLayout();

    import_layout->addWidget(import_button, 0, Qt::AlignCenter);
    import_layout->addWidget(import_progress, 0, Qt::AlignCenter);

    row_input_layout->addWidget(row_label);
    row_input_layout->addWidget(row_input);

    table_layout->addWidget(display);

    main_layout->addLayout(import_layout);
    main_layout->addLayout(row_input_layout);
    main_layout->addLayout(table_layout);
//    display->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);

    setLayout(main_layout);

    // file input signal
    connect(import_button, SIGNAL(clicked()), this, SLOT(importClicked()));
    // row input signal
    connect(row_input, SIGNAL(returnPressed()), this, SLOT(rowEntered()));
}

void Window::importClicked() {
    //Update file_path
    QString temp_file_path = QFileDialog::getOpenFileName(this, tr("Open File"),QDir::currentPath(),tr("Excel Files (*.xlsx)"));
    if(temp_file_path.isEmpty()) return;
    file_path = temp_file_path;

    std::filesystem::path p(file_path.toStdString());
    file_name = QString::fromStdString(p.stem().string());

    //Loading Excel on another thread to prevent UI freezing
    QFuture<void> future = QtConcurrent::run(&Window::loadExcel, this);
}

void Window::loadExcel() {
    //Loading 'animation'
    is_import_loading = true;
    QFuture<void> future = QtConcurrent::run(&Window::importLoadingAnimation, this);

    //Loading Excel into table
    QXlsx::Document doc(file_path);
    if(!doc.load()) {
        is_import_loading = false;
        import_progress->setText("Error Loading " + file_name);
        return;
    }

    int lastColumn = doc.dimension().lastColumn();
    int lastRow = doc.dimension().lastRow();

    table.clear();
    for(int i = 1; i <= lastRow; i++) {
        QVector<QString> data;
        for(int j = 1; j <= lastColumn; j++) {
            QString cell = doc.read(i, j).toString();
            data.push_back(cell);
        }
        table.push_back(data);
    }
    is_import_loading = false;
    row_dimension = table.size();
    col_dimension =  table[0].size();
}

void Window::importLoadingAnimation() {
    import_progress->setText("Importing " + file_name);
    import_progress->show();
    while(is_import_loading) {
       import_progress->setText("Importing " + file_name);
       QThread::currentThread()->msleep(300);
       import_progress->setText("Importing " + file_name + ".");
       QThread::currentThread()->msleep(300);
       import_progress->setText("Importing " + file_name + "..");
       QThread::currentThread()->msleep(300);
       import_progress->setText("Importing " + file_name + "...");
       QThread::currentThread()->msleep(300);
    }
    import_progress->setText("Finished Importing " + file_name);
}

void Window::rowEntered() {
    if(table.empty()) return;
    std::string row_string =  row_input->text().toStdString();
    // input validation
    std::string temp = row_string;
    if(temp == "") return;
    if(temp.find_first_not_of("0123456789, ") != std::string::npos)
    {
       qDebug() << "invalid input: " << row_string;
       QMessageBox row_input_error(this);
       row_input_error.setText("Invalid Row IDs");
       row_input_error.setIcon(QMessageBox::Warning);
       row_input_error.exec();
       return;
    }

    row_ids.clear();
    std::stringstream ss(row_string);
    while(ss.good()) {
       std::string substr;
       getline(ss, substr, ',');
       if(substr =="") continue;
       int id = stoi(substr);
       if(id > row_dimension || id < 1) {
            QMessageBox row_input_error(this);
            row_input_error.setText("Row ID is out of range");
            row_input_error.setIcon(QMessageBox::Warning);
            row_input_error.exec();
            return;
       }
       row_ids.push_back(id);
    }

    if(!(row_ids.size() > 0)) return;
    parseRow();
}

void Window::parseRow() {
    display->setRowCount(col_dimension);
    display->setColumnCount(row_ids.size()+1);
    display->clearContents();

    // Display Headers
    for(int i = 0; i <col_dimension; i++) {
        QTableWidgetItem *item = new QTableWidgetItem(table[0][i]);
        display->setItem(i, 0, item);
    }
    // Display Row Data
    for(int i = 0; i < row_ids.size(); i++) {
        for(int j = 0; j <col_dimension; j++) {
            QTableWidgetItem *item = new QTableWidgetItem(table[row_ids[i]-1][j]);
            display->setItem(j, i+1, item);
        }
    }

    if(display->isHidden()) display->show();
}


