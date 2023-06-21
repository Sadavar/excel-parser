#ifndef WINDOW_H
#define WINDOW_H

#include <QWidget>
//#include "xlsxdocument.h"

class QPushButton;
class QLineEdit;
class QTableWidget;
class QLabel;

class Window : public QWidget {
    Q_OBJECT
public:
    explicit Window(QWidget *parent = 0);
    void parseRow();
    void loadExcel();
     void importLoadingAnimation();
signals:

private slots:
    void importClicked();
    void rowEntered();
private:
    // Widgets
    QPushButton *import_button;
    QLineEdit *row_input;
    QTableWidget *display;
    QLabel *import_progress;
    // Variables
    QString file_path;
    QString file_name;
    QVector<int> row_ids;
    QVector<QVector<QString>> table;
    int row_dimension;
    int col_dimension;
    bool is_import_loading;
};

#endif // WINDOW_H
