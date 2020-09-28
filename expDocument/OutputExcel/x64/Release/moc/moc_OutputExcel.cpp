/****************************************************************************
** Meta object code from reading C++ file 'OutputExcel.h'
**
** Created by: The Qt Meta Object Compiler version 67 (Qt 5.14.0)
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include <memory>
#include "../../../OutputExcel.h"
#include <QtCore/qbytearray.h>
#include <QtCore/qmetatype.h>
#if !defined(Q_MOC_OUTPUT_REVISION)
#error "The header file 'OutputExcel.h' doesn't include <QObject>."
#elif Q_MOC_OUTPUT_REVISION != 67
#error "This file was generated using the moc from 5.14.0. It"
#error "cannot be used with the include files from this version of Qt."
#error "(The moc has changed too much.)"
#endif

QT_BEGIN_MOC_NAMESPACE
QT_WARNING_PUSH
QT_WARNING_DISABLE_DEPRECATED
struct qt_meta_stringdata_OutputExcel_t {
    QByteArrayData data[9];
    char stringdata0[131];
};
#define QT_MOC_LITERAL(idx, ofs, len) \
    Q_STATIC_BYTE_ARRAY_DATA_HEADER_INITIALIZER_WITH_OFFSET(len, \
    qptrdiff(offsetof(qt_meta_stringdata_OutputExcel_t, stringdata0) + ofs \
        - idx * sizeof(QByteArrayData)) \
    )
static const qt_meta_stringdata_OutputExcel_t qt_meta_stringdata_OutputExcel = {
    {
QT_MOC_LITERAL(0, 0, 11), // "OutputExcel"
QT_MOC_LITERAL(1, 12, 12), // "ConfigSignal"
QT_MOC_LITERAL(2, 25, 0), // ""
QT_MOC_LITERAL(3, 26, 22), // "on_BnCreatWord_clicked"
QT_MOC_LITERAL(4, 49, 21), // "on_pushButton_clicked"
QT_MOC_LITERAL(5, 71, 23), // "on_BnConfigJion_clicked"
QT_MOC_LITERAL(6, 95, 10), // "LoadConfig"
QT_MOC_LITERAL(7, 106, 10), // "ShowConfig"
QT_MOC_LITERAL(8, 117, 13) // "hideTabWidget"

    },
    "OutputExcel\0ConfigSignal\0\0"
    "on_BnCreatWord_clicked\0on_pushButton_clicked\0"
    "on_BnConfigJion_clicked\0LoadConfig\0"
    "ShowConfig\0hideTabWidget"
};
#undef QT_MOC_LITERAL

static const uint qt_meta_data_OutputExcel[] = {

 // content:
       8,       // revision
       0,       // classname
       0,    0, // classinfo
       7,   14, // methods
       0,    0, // properties
       0,    0, // enums/sets
       0,    0, // constructors
       0,       // flags
       1,       // signalCount

 // signals: name, argc, parameters, tag, flags
       1,    1,   49,    2, 0x06 /* Public */,

 // slots: name, argc, parameters, tag, flags
       3,    0,   52,    2, 0x08 /* Private */,
       4,    0,   53,    2, 0x08 /* Private */,
       5,    0,   54,    2, 0x08 /* Private */,
       6,    0,   55,    2, 0x08 /* Private */,
       7,    0,   56,    2, 0x08 /* Private */,
       8,    0,   57,    2, 0x08 /* Private */,

 // signals: parameters
    QMetaType::Void, QMetaType::QString,    2,

 // slots: parameters
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,

       0        // eod
};

void OutputExcel::qt_static_metacall(QObject *_o, QMetaObject::Call _c, int _id, void **_a)
{
    if (_c == QMetaObject::InvokeMetaMethod) {
        auto *_t = static_cast<OutputExcel *>(_o);
        Q_UNUSED(_t)
        switch (_id) {
        case 0: _t->ConfigSignal((*reinterpret_cast< QString(*)>(_a[1]))); break;
        case 1: _t->on_BnCreatWord_clicked(); break;
        case 2: _t->on_pushButton_clicked(); break;
        case 3: _t->on_BnConfigJion_clicked(); break;
        case 4: _t->LoadConfig(); break;
        case 5: _t->ShowConfig(); break;
        case 6: _t->hideTabWidget(); break;
        default: ;
        }
    } else if (_c == QMetaObject::IndexOfMethod) {
        int *result = reinterpret_cast<int *>(_a[0]);
        {
            using _t = void (OutputExcel::*)(QString );
            if (*reinterpret_cast<_t *>(_a[1]) == static_cast<_t>(&OutputExcel::ConfigSignal)) {
                *result = 0;
                return;
            }
        }
    }
}

QT_INIT_METAOBJECT const QMetaObject OutputExcel::staticMetaObject = { {
    QMetaObject::SuperData::link<QMainWindow::staticMetaObject>(),
    qt_meta_stringdata_OutputExcel.data,
    qt_meta_data_OutputExcel,
    qt_static_metacall,
    nullptr,
    nullptr
} };


const QMetaObject *OutputExcel::metaObject() const
{
    return QObject::d_ptr->metaObject ? QObject::d_ptr->dynamicMetaObject() : &staticMetaObject;
}

void *OutputExcel::qt_metacast(const char *_clname)
{
    if (!_clname) return nullptr;
    if (!strcmp(_clname, qt_meta_stringdata_OutputExcel.stringdata0))
        return static_cast<void*>(this);
    return QMainWindow::qt_metacast(_clname);
}

int OutputExcel::qt_metacall(QMetaObject::Call _c, int _id, void **_a)
{
    _id = QMainWindow::qt_metacall(_c, _id, _a);
    if (_id < 0)
        return _id;
    if (_c == QMetaObject::InvokeMetaMethod) {
        if (_id < 7)
            qt_static_metacall(this, _c, _id, _a);
        _id -= 7;
    } else if (_c == QMetaObject::RegisterMethodArgumentMetaType) {
        if (_id < 7)
            *reinterpret_cast<int*>(_a[0]) = -1;
        _id -= 7;
    }
    return _id;
}

// SIGNAL 0
void OutputExcel::ConfigSignal(QString _t1)
{
    void *_a[] = { nullptr, const_cast<void*>(reinterpret_cast<const void*>(std::addressof(_t1))) };
    QMetaObject::activate(this, &staticMetaObject, 0, _a);
}
QT_WARNING_POP
QT_END_MOC_NAMESPACE
