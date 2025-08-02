package main

import (
	"fmt"
	"image"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/gen2brain/go-fitz" // Untuk render thumbnail
	"github.com/unidoc/unipdf/v3/common/license"
	"github.com/unidoc/unipdf/v3/creator"
	pdf "github.com/unidoc/unipdf/v3/model"
)

const appName = "PDF Merger Offline"

type PDFFile struct {
	Path     string
	PageCount int
	Preview  image.Image
}

func main() {
	// Setup aplikasi
	myApp := app.New()
	myWindow := myApp.NewWindow(appName)
	myWindow.Resize(fyne.NewSize(900, 600))

	// Inisialisasi state
	var pdfFiles []PDFFile
	var selectedIndex int = -1
	statusLabel := widget.NewLabel("Ready")
	previewImage := canvas.NewImageFromImage(nil)
	previewImage.FillMode = canvas.ImageFillOriginal
	outputEntry := widget.NewEntry()
	outputEntry.SetText("merged.pdf")

	// Komponen UI
	fileList := widget.NewList(
		func() int { return len(pdfFiles) },
		func() fyne.CanvasObject {
			return container.NewHBox(
				widget.NewLabel(""),
				layout.NewSpacer(),
				widget.NewLabel(""),
			)
		},
		func(i widget.ListItemID, o fyne.CanvasObject) {
			cont := o.(*fyne.Container)
			labels := cont.Objects
			file := pdfFiles[i]
			labels[0].(*widget.Label).SetText(fmt.Sprintf("%d. %s", i+1, filepath.Base(file.Path)))
			labels[2].(*widget.Label).SetText(fmt.Sprintf("%d pages", file.PageCount))
		},
	)

	// Fungsi untuk update preview
	updatePreview := func() {
		if selectedIndex >= 0 && selectedIndex < len(pdfFiles) {
			if pdfFiles[selectedIndex].Preview != nil {
				previewImage.Image = pdfFiles[selectedIndex].Preview
				previewImage.Refresh()
			}
		} else {
			previewImage.Image = nil
			previewImage.Refresh()
		}
	}

	// Event handler untuk seleksi file
	fileList.OnSelected = func(id widget.ListItemID) {
		selectedIndex = int(id)
		updatePreview()
	}

	// Fungsi untuk menambah file PDF
	addPDFFiles := func(paths []string) {
		for _, path := range paths {
			// Cek apakah file sudah ada
			exists := false
			for _, f := range pdfFiles {
				if f.Path == path {
					exists = true
					break
				}
			}
			
			if exists {
				continue
			}

			// Hitung jumlah halaman
			pageCount, err := getPDFPageCount(path)
			if err != nil {
				dialog.ShowError(fmt.Errorf("Error reading %s: %v", filepath.Base(path), err), myWindow)
				continue
			}

			// Buat preview thumbnail
			preview, _ := renderPDFPreview(path)

			pdfFiles = append(pdfFiles, PDFFile{
				Path:      path,
				PageCount: pageCount,
				Preview:   preview,
			})
		}
		fileList.Refresh()
		statusLabel.SetText(fmt.Sprintf("Added %d files. Total: %d", len(paths), len(pdfFiles)))
	}

	// Action buttons
	moveUpBtn := widget.NewButton("Move Up", func() {
		if selectedIndex > 0 {
			pdfFiles[selectedIndex], pdfFiles[selectedIndex-1] = pdfFiles[selectedIndex-1], pdfFiles[selectedIndex]
			selectedIndex--
			fileList.Refresh()
			fileList.Select(selectedIndex)
		}
	})

	moveDownBtn := widget.NewButton("Move Down", func() {
		if selectedIndex >= 0 && selectedIndex < len(pdfFiles)-1 {
			pdfFiles[selectedIndex], pdfFiles[selectedIndex+1] = pdfFiles[selectedIndex+1], pdfFiles[selectedIndex]
			selectedIndex++
			fileList.Refresh()
			fileList.Select(selectedIndex)
		}
	})

	removeBtn := widget.NewButton("Remove Selected", func() {
		if selectedIndex >= 0 && selectedIndex < len(pdfFiles) {
			// Hapus file dari slice
			pdfFiles = append(pdfFiles[:selectedIndex], pdfFiles[selectedIndex+1:]...)
			
			// Reset seleksi
			if len(pdfFiles) > 0 {
				if selectedIndex >= len(pdfFiles) {
					selectedIndex = len(pdfFiles) - 1
				}
				fileList.Select(selectedIndex)
			} else {
				selectedIndex = -1
				previewImage.Image = nil
				previewImage.Refresh()
			}
			
			fileList.Refresh()
			statusLabel.SetText("Removed 1 file")
		}
	})

	browseOutputBtn := widget.NewButton("Browse", func() {
		dialog.ShowFileSave(func(uri fyne.URIWriteCloser, err error) {
			if err != nil || uri == nil {
				return
			}
			outputPath := uri.URI().Path()
			if !strings.HasSuffix(outputPath, ".pdf") {
				outputPath += ".pdf"
			}
			outputEntry.SetText(outputPath)
			uri.Close()
		}, myWindow)
	})

	mergeBtn := widget.NewButton("Merge PDF", func() {
		if len(pdfFiles) == 0 {
			dialog.ShowInformation("Warning", "No PDF files to merge!", myWindow)
			return
		}

		outputPath := outputEntry.Text
		if outputPath == "" {
			dialog.ShowInformation("Warning", "Please specify output file path!", myWindow)
			return
		}

		// Tampilkan progress dialog
		progress := widget.NewProgressBarInfinite()
		progressDialog := dialog.NewCustom("Merging PDFs", "Cancel", progress, myWindow)
		progressDialog.Show()

		// Jalankan merge di goroutine terpisah
		go func() {
			err := mergePDFs(pdfFiles, outputPath)
			progressDialog.Hide()
			
			if err != nil {
				dialog.ShowError(fmt.Errorf("Merge failed: %v", err), myWindow)
				statusLabel.SetText("Merge failed")
			} else {
				dialog.ShowInformation("Success", "PDF files merged successfully!", myWindow)
				statusLabel.SetText(fmt.Sprintf("Created: %s", filepath.Base(outputPath)))
			}
		}()
	})

	// Layout UI
	previewContainer := container.NewVBox(
		widget.NewLabel("Preview (first page)"),
		container.NewCenter(previewImage),
	)

	outputContainer := container.NewBorder(nil, nil, widget.NewLabel("Output File:"), nil, outputEntry)
	buttonContainer := container.NewHBox(
		widget.NewButton("Select Files", func() {
			dialog.ShowFileOpen(func(uris []fyne.URIReadCloser, err error) {
				if err != nil || uris == nil {
					return
				}
				
				var paths []string
				for _, uri := range uris {
					paths = append(paths, uri.URI().Path())
					uri.Close()
				}
				addPDFFiles(paths)
			}, myWindow)
		}),
		moveUpBtn,
		moveDownBtn,
		removeBtn,
		browseOutputBtn,
		mergeBtn,
	)

	mainContainer := container.NewBorder(
		nil, 
		container.NewVBox(buttonContainer, statusLabel), 
		nil, 
		nil,
		container.NewHSplit(
			container.NewBorder(
				nil, 
				nil, 
				nil, 
				nil, 
				fileList,
			),
			previewContainer,
		),
	)

	// Setup drag and drop
	myWindow.SetContent(mainContainer)
	myWindow.SetOnDropped(func(pos fyne.Position, uris []fyne.URI) {
		var paths []string
		for _, uri := range uris {
			path := uri.Path()
			if strings.ToLower(filepath.Ext(path)) == ".pdf" {
				paths = append(paths, path)
			}
		}
		if len(paths) > 0 {
			addPDFFiles(paths)
		}
	})

	// Init unipdf license (gratis untuk non-komersial)
	license.SetMeteredKey("") // Isi dengan API key jika punya

	myWindow.ShowAndRun()
}

// Render preview halaman pertama PDF
func renderPDFPreview(pdfPath string) (image.Image, error) {
	doc, err := fitz.New(pdfPath)
	if err != nil {
		return nil, err
	}
	defer doc.Close()

	// Render halaman pertama
	img, err := doc.Image(0)
	if err != nil {
		return nil, err
	}

	return img, nil
}

// Mendapatkan jumlah halaman PDF
func getPDFPageCount(pdfPath string) (int, error) {
	doc, err := fitz.New(pdfPath)
	if err != nil {
		return 0, err
	}
	defer doc.Close()

	return doc.NumPage(), nil
}

// Merge multiple PDF files
func mergePDFs(files []PDFFile, outputPath string) error {
	c := creator.New()

	for _, file := range files {
		// Baca file PDF
		f, err := os.Open(file.Path)
		if err != nil {
			return err
		}
		defer f.Close()

		pdfReader, err := pdf.NewPdfReader(f)
		if err != nil {
			return err
		}

		numPages, err := pdfReader.GetNumPages()
		if err != nil {
			return err
		}

		// Tambahkan semua halaman ke creator
		for i := 0; i < numPages; i++ {
			pageNum := i + 1

			page, err := pdfReader.GetPage(pageNum)
			if err != nil {
				return err
			}

			if err := c.AddPage(page); err != nil {
				return err
			}
		}
	}

	// Simpan hasil merge
	return c.WriteToFile(outputPath)
}