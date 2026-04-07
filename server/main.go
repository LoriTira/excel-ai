package main

import (
	"crypto/tls"
	"embed"
	"fmt"
	"io/fs"
	"net/http"
	"net/http/httputil"
	"net/url"
	"os"
	"path/filepath"
)

//go:embed static
var staticFiles embed.FS

func main() {
	home, err := os.UserHomeDir()
	if err != nil {
		fatal("cannot find home directory: %v", err)
	}

	certFile := filepath.Join(home, ".ollama", "cert.pem")
	keyFile := filepath.Join(home, ".ollama", "key.pem")

	// Reverse proxy to Ollama
	target, _ := url.Parse("http://127.0.0.1:11434")
	proxy := httputil.NewSingleHostReverseProxy(target)

	// Static file server from embedded dist files
	sub, _ := fs.Sub(staticFiles, "static")
	files := http.FileServer(http.FS(sub))

	mux := http.NewServeMux()
	mux.HandleFunc("/v1/", proxy.ServeHTTP)
	mux.HandleFunc("/api/", proxy.ServeHTTP)
	mux.Handle("/", files)

	cert, err := tls.LoadX509KeyPair(certFile, keyFile)
	if err != nil {
		fatal("TLS cert error: %v\nRun the install script to generate certificates.", err)
	}

	srv := &http.Server{
		Addr:    "127.0.0.1:11435",
		Handler: mux,
		TLSConfig: &tls.Config{
			Certificates: []tls.Certificate{cert},
		},
	}

	fmt.Println("Excel AI: https://127.0.0.1:11435 -> Ollama http://127.0.0.1:11434")
	if err := srv.ListenAndServeTLS("", ""); err != nil {
		fatal("server: %v", err)
	}
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}
