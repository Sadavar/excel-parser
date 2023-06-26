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
#include <QSpacerItem>
#include <QListWidget>
#include <QListWidgetItem>


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
    import_button = new QPushButton("  Import File");
    import_button->setObjectName("import_button");
    import_button->setCheckable(true);
    import_button->setFixedHeight(50);
    import_button->setMinimumWidth(250);
    import_button->setMaximumWidth(550);
    import_button->setIcon(QIcon(":/upload_icon.png"));

    // Create import progress label
    import_progress = new QLabel("Importing...");
    import_progress->hide();

    // Create row input field
    auto row_label = new QLabel("Row IDs:");
    row_input = new QLineEdit();
    row_input->setPlaceholderText("2,5,8");

    // Create project column input
    auto project_column_input_label = new QLabel("Project Column Letter:");
    project_column_input = new QLineEdit();
    project_column_input->setPlaceholderText("5");
    //Create project number input
    auto project_number_input_label = new QLabel("Project Number:");
    project_number_input = new QLineEdit();
    project_number_input->setPlaceholderText("A21-1639");

    //Create filter
    filter = new QListWidget();
    filter->setSelectionMode(QAbstractItemView::MultiSelection);

    // Create Row ID Table
    row_display = new QTableWidget();
    row_display->setVerticalScrollMode(QAbstractItemView::ScrollPerPixel);
    row_display->setHorizontalScrollMode(QAbstractItemView::ScrollPerPixel);
    row_display_label = new QLabel("Row Search");
    row_display_label->setObjectName("row_display_label");

    // Create Project Table
    project_display = new QTableWidget();
    project_display->setVerticalScrollMode(QAbstractItemView::ScrollPerPixel);
    project_display->setHorizontalScrollMode(QAbstractItemView::ScrollPerPixel);
    project_display_label = new QLabel("Project Search");
    project_display_label->setObjectName("project_display_label");

    //Creating layouts
    main_layout = new QVBoxLayout(this);
    auto import_layout = new QVBoxLayout();
    auto row_input_layout = new QHBoxLayout();
    auto project_input_layout = new QHBoxLayout();
    auto table_layout = new QHBoxLayout();
    auto row_display_layout = new QVBoxLayout();
    auto project_display_layout = new QVBoxLayout();

    import_layout->addWidget(import_button, 0, Qt::AlignCenter);
    import_layout->addWidget(import_progress, 0, Qt::AlignCenter);

    row_input_layout->addWidget(row_label);
    row_input_layout->addWidget(row_input, Qt::AlignTop);

    project_input_layout->addWidget(project_column_input_label);
    project_input_layout->addWidget(project_column_input);
    project_input_layout->addWidget(project_number_input_label);
    project_input_layout->addWidget(project_number_input);

    row_display_layout->addWidget(row_display_label, 0, Qt::AlignCenter);
    row_display_layout->addWidget(row_display);

    project_display_layout->addWidget(project_display_label, 0, Qt::AlignCenter);
    project_display_layout->addWidget(project_display);

    table_layout->addLayout(row_display_layout);
    table_layout->addLayout(project_display_layout);

    main_layout->addLayout(import_layout);
    main_layout->addLayout(row_input_layout);
    main_layout->addLayout(project_input_layout);
    main_layout->addWidget(filter);
    spacer = new QSpacerItem(30, 30, QSizePolicy::Expanding, QSizePolicy::Fixed);
    main_layout->addItem(spacer);
    main_layout->addLayout(table_layout);


    setLayout(main_layout);

    // file input signal
    connect(import_button, SIGNAL(clicked()), this, SLOT(importClicked()));
    // row input signal
    connect(row_input, SIGNAL(returnPressed()), this, SLOT(rowEntered()));
    // project column input signal
    connect(project_column_input, SIGNAL(returnPressed()), this, SLOT(projectColumnEntered()));
    // project number input signal
    connect(project_number_input, SIGNAL(returnPressed()), this, SLOT(projectNumberEntered()));

    has_entered_column = false;
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
            if(i == 1) {
                header_list.append(cell);
            }
        }
        table.push_back(data);
    }
    // add headers to filter
    filter->addItems(header_list);

    is_import_loading = false;
    row_dimension = table.size();
    col_dimension =  table[0].size();
}

void Window::importLoadingAnimation() {
    import_progress->setText("Importing " + file_name);
    import_progress->show();
    this->update();
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
    this->update();
}

void Window::rowEntered() {
    if(table.empty()) {
       if(project_display->isHidden()) {
            spacer->changeSize(10,10, QSizePolicy::Expanding, QSizePolicy::Expanding);
       }
       row_display->hide();
       row_display_label->hide();
       return;
    }
    std::string row_string =  row_input->text().toStdString();
    // input validation
    std::string temp = row_string;
    if(temp.empty()) {
       if(project_display->isHidden()) {
            spacer->changeSize(10,10, QSizePolicy::Expanding, QSizePolicy::Expanding);
       }
       row_display->hide();
       row_display_label->hide();
       return;
    }
    if(temp.find_first_not_of("0123456789, ") != std::string::npos)
    {
       qDebug() << "invalid input: " << row_string;
       QMessageBox row_input_error(this);
       row_input_error.setText("Invalid row input");
       row_input_error.setIcon(QMessageBox::Warning);
       row_input_error.setStandardButtons(QMessageBox::Ok);
       row_input_error.exec();
       return;
    }

    // Parsing input and adding it to row_ids vector
    row_ids.clear();
    std::stringstream ss(row_string);
    while(ss.good()) {
       std::string substr;
       getline(ss, substr, ',');
       //check if the substring is just spaces (eg. 2,5, )
       if(substr.find_first_not_of(' ') == std::string::npos) continue;
       int id = stoi(substr);
       if(id > row_dimension || id < 1) {
            QMessageBox row_input_error(this);
            row_input_error.setText("Row input is out of range");
            row_input_error.setIcon(QMessageBox::Warning);
            row_input_error.setStandardButtons(QMessageBox::Ok);
            row_input_error.exec();
            return;
       }
       row_ids.push_back(id);
    }

    if(!(row_ids.size() > 0)) return;
    parseRow();
}

void Window::parseRow() {
    row_display->setRowCount(col_dimension);
    row_display->setColumnCount(row_ids.size()+1);
    row_display->clearContents();

    // Display Headers
    for(int i = 0; i <col_dimension; i++) {
        QTableWidgetItem *item = new QTableWidgetItem(table[0][i]);
        row_display->setItem(i, 0, item);
    }
    // Display Row Data
    for(int i = 0; i < row_ids.size(); i++) {
        for(int j = 0; j <col_dimension; j++) {
            QTableWidgetItem *item = new QTableWidgetItem(table[row_ids[i]-1][j]);
            row_display->setItem(j, i+1, item);
        }
    }

    spacer->changeSize(30, 30, QSizePolicy::Fixed, QSizePolicy::Fixed);

    if(row_display->isHidden()) {
        row_display->show();
        row_display_label->show();
    }
}

void Window::projectColumnEntered() {
    std::string input =  project_column_input->text().toStdString();
    if(input.empty()) return;
    if(input.find_first_not_of("ABCDEFGHIJKLMNOPQRSTUVWXYZ") != std::string::npos) {
        QMessageBox error(this);
        error.setText("Invalid project column input");
        error.setIcon(QMessageBox::Warning);
        error.setStandardButtons(QMessageBox::Ok);
        error.exec();
        return;
    }

    // Convert column letter to column number (e.g. AB = 27)
    const char *colstr= input.c_str();
    int i, col=0;
    for(i=0; i< (int)input.size(); i++) {
        col = 26*col + colstr[i] - 'A' + 1;
    }
    project_column_number = col;
    has_entered_column = true;
}

void Window::projectNumberEntered() {
    if(table.empty() || !has_entered_column) {
        if(row_display->isHidden()) {
            spacer->changeSize(10,10, QSizePolicy::Expanding, QSizePolicy::Expanding);
        }
        project_display->hide();
        project_display_label->hide();
        return;
    }
    QString input =  project_number_input->text();
    if(input.isEmpty()) {
        if(row_display->isHidden()) {
            spacer->changeSize(10,10, QSizePolicy::Expanding, QSizePolicy::Expanding);
        }
        project_display->hide();
        project_display_label->hide();
        return;
    }

    // Delete trailing whitespace
    std::string input_string = input.toStdString();
    const char* WhiteSpace = " \t\v\r\n";
    std::size_t end = input_string.find_last_not_of(WhiteSpace);
    input_string = input_string.substr(0, end+1);
    qDebug() << "original: " << input;
    qDebug() << "after: " << input_string;

    project_number = QString::fromStdString(input_string);
    parseProject();
}

void Window::parseProject() {
    project_display->setRowCount(col_dimension);
    project_display->setColumnCount(row_dimension);
    project_display->clearContents();

    // Display headers
    for(int i = 0; i < col_dimension; i++) {
        QTableWidgetItem *item = new QTableWidgetItem(table[0][i]);
        project_display->setItem(i, 0, item);
    }

    // Display matching project numbers
    int match_counter = 0;
    int col = project_column_number - 1;
    for(int i =0; i < row_dimension; i++) {
        QString val = table[i][col];
        if(val == project_number) {
            for(int j = 0; j < col_dimension; j++) {
                QTableWidgetItem *item = new QTableWidgetItem(table[i][j]);
                project_display->setItem(j, match_counter + 1, item);
            }
            ++match_counter;
        }
    }
    project_display->setColumnCount(match_counter + 1);
    spacer->changeSize(30, 30, QSizePolicy::Expanding, QSizePolicy::Fixed);
    if(project_display->isHidden()) {
        project_display->show();
        project_display_label->show();
    }
}


