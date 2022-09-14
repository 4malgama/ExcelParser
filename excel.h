#ifndef EXCEL_H
#define EXCEL_H
#include <QAxObject>

class QString;

class Excel : public QObject
{
    QAxObject *ax_excel = nullptr;
    QAxObject *ax_books;
    QAxObject *ax_book;
    QAxObject *ax_sheet;
    QAxObject *ax_range;

    QAxObject *ax_rows;
    QAxObject *ax_cols;

public:
    Excel(const QString &path);
    ~Excel();

    int rows() const;
    int columns() const;
    int firstRow() const;
    int firstCol() const;
    int lastRow() const;
    int lastCol() const;
    int column(const QString &address);
    int row(const QString &address);

    QVariant cellData(const int &row, const int &col);
    QVariant cellData(const QString &address);

    QString author() const;
    QString name() const;
    QString tableName() const;
    QString findCellByText(const QString &text);
    QString cellAddress(const int &row, const int &col) const;

    QList<QString> findCellsByText(const QString &text);

    void createDoc(QAxObject *object, const QString &name = "");
    void setTable(const int &index);

    QVariant call_function(QAxObject*, const char*, QList<QVariant> &);

    bool isNull();

private:
    int _rows;
    int _cols;

    int _firstRow;
    int _firstCol;

    int _lastRow;
    int _lastCol;

    void init(int index);
    void finish();
};

#endif // EXCEL_H
