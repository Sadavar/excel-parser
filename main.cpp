#include <QApplication>
#include "window.h"
#include <QFile>

int main(int argc, char **argv) {
    QApplication app (argc, argv);

    // Load an application style
    QFile styleFile( ":/style.qss" );
    styleFile.open( QFile::ReadOnly );

    // Apply the loaded stylesheet
    QString style( styleFile.readAll() );
    app.setStyleSheet( style );

    Window window;
    window.show();

    return app.exec();
}
