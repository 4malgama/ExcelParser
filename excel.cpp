#include "excel.h"

#include <QDebug>

#define _VAR(x) #x

#define DOCUMENTATION

#ifdef DOCUMENTATION
#include <QFile>
#include <QDir>
#endif

class __Reader
{
    QByteArray data;
    char c;
    int pos;

public:
    __Reader(const QByteArray &data, char c = ':') : data{data}, c{c}, pos{} {}

    QByteArray readArgument()
    {
        QByteArray arg;
        int i = data.indexOf(c, pos);
        if (i == -1) return readRest();
        arg.setRawData(data.constData() + pos, i - pos);
        pos = i + 1;
        return arg;
    }

    QByteArray readLast()
    {
        int i = data.lastIndexOf(c);
        if (i == -1) return readRest();
        i++;
        return QByteArray::fromRawData(data.constData() + i, data.size() - i);
    }

    QByteArray readRest()
    {
        return QByteArray::fromRawData(data.constData() + pos, data.size() - pos);
    }

    QByteArray& readAll()
    {
        return data;
    }

    void setC(char c) { this->c = c; }
};

Excel::Excel(const QString &path)
{
    ax_excel = new QAxObject("Excel.Application", this);

    ax_books = ax_excel->querySubObject("Workbooks");
    ax_book = ax_books->querySubObject("Open(const QString&)", path);
    ax_excel->dynamicCall("SetVisible(bool)", false);

    init(1);

#ifdef DOCUMENTATION
    createDoc(ax_excel, "ax_excel");
    createDoc(ax_books, "ax_books");
    createDoc(ax_book, "ax_book");
    createDoc(ax_sheet, "ax_sheet");
    createDoc(ax_range, "ax_range");
    createDoc(ax_rows, "ax_rows");
    createDoc(ax_cols, "ax_cols");
#endif
}

Excel::~Excel()
{
    finish();
}

int Excel::rows() const
{
    return this->_rows;
}

int Excel::columns() const
{
    return this->_cols;
}

int Excel::firstRow() const
{
    return this->_firstRow;
}

int Excel::firstCol() const
{
    return this->_firstCol;
}

int Excel::lastRow() const
{
    return this->_lastRow;
}

int Excel::lastCol() const
{
    return this->_lastCol;
}

int Excel::column(const QString &address)
{
    QAxObject *cell = ax_sheet->querySubObject("Range(const QString&)", address);
    return cell->property("Column").toInt();
}

int Excel::row(const QString &address)
{
    QAxObject *cell = ax_sheet->querySubObject("Range(const QString&)", address);
    return cell->property("Row").toInt();
}

QVariant Excel::cellData(const int &row, const int &col)
{
    QAxObject *cell = ax_sheet->querySubObject("Cells(int,int)", row, col);
    QVariant value = cell->dynamicCall("Value()");
    return value;
}

QVariant Excel::cellData(const QString &address)
{
    QAxObject *cell = ax_sheet->querySubObject("Range(const QString&)", address);
    QVariant value = cell->dynamicCall("Value()");
    return value;
}

QString Excel::author() const
{
    return ax_book->property("Author").toString();
}

QString Excel::name() const
{
    return ax_book->property("Name").toString();
}

QString Excel::tableName() const
{
    return ax_sheet->property("Name").toString();
}

QString Excel::findCellByText(const QString &text)
{
    for (int i{_firstRow}; i <= _lastRow; i++)
    {
        for (int r{_firstCol}; r <= _lastCol; r++)
        {
            QAxObject *cell = ax_sheet->querySubObject("Cells(int,int)", i, r);
            if (cell->dynamicCall("Value()").toString() == text)
            {
                return cell->dynamicCall("Address()").toString();
            }
        }
    }
    return "";
}

QString Excel::cellAddress(const int &row, const int &col) const
{
    QAxObject *cell = ax_sheet->querySubObject("Cells(int,int)", row, col);
    return cell->dynamicCall("Address()").toString();
}

QList<QString> Excel::findCellsByText(const QString &text)
{
    QList<QString> list;
    for (int i{_firstRow}; i <= _lastRow; i++)
    {
        for (int r{_firstCol}; r <= _lastCol; r++)
        {
            QAxObject *cell = ax_sheet->querySubObject("Cells(int,int)", i, r);
            if (cell->dynamicCall("Value()").toString() == text)
            {
                list.append(cell->dynamicCall("Address()").toString());
            }
        }
    }
    return list;
}

void Excel::createDoc(QAxObject *object, const QString &name)
{
#ifdef DOCUMENTATION
{
    QByteArray doc = object->generateDocumentation().toUtf8();
    QDir(QDir::currentPath()).mkdir("all_interfaces");
    QFile file(QDir::currentPath() + "\\all_interfaces\\interface_" + name + ".html");
    if (file.open(QFile::WriteOnly))
    {
        file.write(doc);
        file.close();
    }
    else
    {
        qWarning() << "Error: Document file not created - " << name;
    }
}
#endif
}

void Excel::setTable(const int &index)
{
    init(index);
}

bool Excel::isNull()
{
    return (ax_excel == nullptr);
}

void Excel::init(int index)
{
    ax_sheet = ax_book->querySubObject("WorkSheets(int)", index);

    ax_range = ax_sheet->querySubObject("UsedRange");

    ax_rows = ax_range->querySubObject("Rows");
    ax_cols = ax_range->querySubObject("Columns");

    _cols = ax_cols->property("Count").toInt();
    _rows = ax_rows->property("Count").toInt();
    _firstCol = ax_range->property("Column").toInt();
    _firstRow = ax_range->property("Row").toInt();
    {
        QByteArray range = ax_range->dynamicCall("Address(False, False)").toByteArray();

        __Reader R(range);
        _lastRow = row(R.readLast());
        _lastCol = column(R.readLast());
    }
}

void Excel::finish()
{
    ax_book->dynamicCall("Close");
    ax_excel->dynamicCall("Quit()");

    ax_cols->clear();
    ax_rows->clear();
    ax_range->clear();
    ax_sheet->clear();
    ax_book->clear();
    ax_books->clear();
    ax_excel->clear();

    delete ax_excel;
    ax_excel = nullptr;
}

QVariant Excel::call_function(QAxObject *object, const char *function, QList<QVariant> &arguments)
{
    return object->dynamicCall(function, arguments);
}
