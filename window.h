#ifndef WINDOW_H
#define WINDOW_H

#include <QWidget>
//#include "xlsxdocument.h"

class QPushButton;
class QLineEdit;
class QTableWidget;
class QLabel;
class QVBoxLayout;
class QSpacerItem;
class QListWidget;

class Window : public QWidget {
    Q_OBJECT
public:
    explicit Window(QWidget *parent = 0);
    void parseRow();
    void parseProject();
    void loadExcel();
     void importLoadingAnimation();
signals:

private slots:
    void importClicked();
    void rowEntered();
    void projectColumnEntered();
    void projectNumberEntered();
private:
    // Widgets
    QPushButton *import_button;
    QLineEdit *row_input;
    QLineEdit *project_column_input;
    QLineEdit *project_number_input;
    QTableWidget *row_display;
    QLabel *row_display_label;
    QTableWidget *project_display;
    QLabel *project_display_label;
    QLabel *import_progress;
    QVBoxLayout *main_layout;
    QSpacerItem *spacer;
    QListWidget *filter;
    // Variables
    QString file_path;
    QString file_name;
    QVector<int> row_ids;
    QVector<QVector<QString>> table;
    int row_dimension;
    int col_dimension;
    bool is_import_loading;
    bool has_entered_column;
    int project_column_number;
    QString project_number;
    QStringList header_list;
};

#endif // WINDOW_H
