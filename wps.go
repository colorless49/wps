// +build windows

package main

import (
	"fmt"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const wdFormatPDF = 17

func main() {
	docFile := "d:/test.doc"
	pdfFile := "d:/test.pdf"

	ole.CoInitialize(0)
	unknown, _ := oleutil.CreateObject("KWPS.Application")
	kwps, _ := unknown.QueryInterface(ole.IID_IDispatch)
	oleutil.PutProperty(kwps, "Visible", true)
	fmt.Printf("wps版本：" + oleutil.MustGetProperty(kwps, "Version").ToString())

	docs := oleutil.MustGetProperty(kwps, "Documents").ToIDispatch()

	doc := oleutil.MustCallMethod(docs, "Open", docFile, false, true).ToIDispatch()
	//Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, wdFormatPDF );

	oleutil.MustCallMethod(doc, "ExportAsFixedFormat", pdfFile, wdFormatPDF)
	//Dispatch.call(doc, "Close", false);
	oleutil.MustCallMethod(doc, "Close", false)
	doc.Release()
	docs.Release()
	oleutil.CallMethod(kwps, "Quit")
	kwps.Release()

	ole.CoUninitialize()
}
