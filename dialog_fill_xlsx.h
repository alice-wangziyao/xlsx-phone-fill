#ifndef DIALOG_FILL_XLSX_H
#define DIALOG_FILL_XLSX_H

#include <QDialog>

namespace Ui {
class dialog_fill_xlsx;
}

class dialog_fill_xlsx : public QDialog
{
    Q_OBJECT

public:
    explicit dialog_fill_xlsx(QWidget *parent = nullptr);
    ~dialog_fill_xlsx();

private:
    Ui::dialog_fill_xlsx *ui;
};

#endif // DIALOG_FILL_XLSX_H
