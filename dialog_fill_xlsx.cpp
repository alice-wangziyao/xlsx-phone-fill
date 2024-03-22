#include "dialog_fill_xlsx.h"
#include "ui_dialog_fill_xlsx.h"

dialog_fill_xlsx::dialog_fill_xlsx(QWidget *parent)
    : QDialog(parent)
    , ui(new Ui::dialog_fill_xlsx)
{
    ui->setupUi(this);
}

dialog_fill_xlsx::~dialog_fill_xlsx()
{
    delete ui;
}
