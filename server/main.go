package main

import (
	"crypto/tls"
	"embed"
	"encoding/json"
	"fmt"
	"io/fs"
	"net/http"
	"net/http/httputil"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sync"
	"time"
)

//go:embed static
var staticFiles embed.FS

const (
	ollamaURL        = "http://127.0.0.1:11434"
	listenAddr       = "127.0.0.1:11435"
	heartbeatTimeout = 60 * time.Second
	watchdogInterval = 10 * time.Second
	startPollDelay   = 500 * time.Millisecond
	startTimeout     = 10 * time.Second
)

// Ollama lifecycle state
var (
	mu             sync.Mutex
	ollamaCmd      *exec.Cmd
	ollamaRunning  bool
	ollamaStarted  bool // true if we spawned the process
	lastHeartbeat  time.Time
	ollamaStartedAt time.Time
)

func startOllama() error {
	mu.Lock()
	defer mu.Unlock()

	if ollamaRunning {
		return nil
	}

	// Check if Ollama is already running externally
	if isOllamaReachable() {
		ollamaRunning = true
		ollamaStarted = false
		ollamaStartedAt = time.Now()
		go watchdog()
		return nil
	}

	// Spawn Ollama
	cmd := exec.Command("ollama", "serve")
	cmd.Stdout = nil
	cmd.Stderr = nil
	if err := cmd.Start(); err != nil {
		return fmt.Errorf("failed to start ollama: %w", err)
	}

	ollamaCmd = cmd
	ollamaRunning = true
	ollamaStarted = true
	ollamaStartedAt = time.Now()

	// Wait for Ollama to become reachable (unlock while polling)
	mu.Unlock()
	err := waitForOllama()
	mu.Lock()

	if err != nil {
		// Failed to start — clean up
		ollamaCmd.Process.Kill()
		ollamaCmd.Wait()
		ollamaCmd = nil
		ollamaRunning = false
		ollamaStarted = false
		return err
	}

	go watchdog()
	return nil
}

func stopOllama() {
	mu.Lock()
	defer mu.Unlock()

	if !ollamaRunning {
		return
	}

	if ollamaStarted && ollamaCmd != nil && ollamaCmd.Process != nil {
		if runtime.GOOS == "windows" {
			ollamaCmd.Process.Kill()
		} else {
			ollamaCmd.Process.Signal(os.Interrupt)
		}
		ollamaCmd.Wait()
	}

	ollamaCmd = nil
	ollamaRunning = false
	ollamaStarted = false
	fmt.Println("Ollama stopped.")
}

func watchdog() {
	for {
		time.Sleep(watchdogInterval)

		mu.Lock()
		if !ollamaRunning {
			mu.Unlock()
			return
		}
		elapsed := time.Since(lastHeartbeat)
		mu.Unlock()

		if elapsed > heartbeatTimeout {
			fmt.Println("No heartbeat — stopping Ollama.")
			stopOllama()
			return
		}
	}
}

func isOllamaReachable() bool {
	client := &http.Client{Timeout: time.Second}
	resp, err := client.Get(ollamaURL)
	if err != nil {
		return false
	}
	resp.Body.Close()
	return true
}

func waitForOllama() error {
	deadline := time.Now().Add(startTimeout)
	for time.Now().Before(deadline) {
		if isOllamaReachable() {
			return nil
		}
		time.Sleep(startPollDelay)
	}
	return fmt.Errorf("ollama did not start within %s", startTimeout)
}

func ensureOllama(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		mu.Lock()
		running := ollamaRunning
		mu.Unlock()

		if !running {
			if err := startOllama(); err != nil {
				http.Error(w, fmt.Sprintf("Failed to start Ollama: %v", err), http.StatusBadGateway)
				return
			}
		}
		next(w, r)
	}
}

type heartbeatResponse struct {
	Ollama string  `json:"ollama"`
	Uptime float64 `json:"uptime,omitempty"`
}

func handleHeartbeat(w http.ResponseWriter, r *http.Request) {
	mu.Lock()
	lastHeartbeat = time.Now()
	running := ollamaRunning
	mu.Unlock()

	if !running {
		if err := startOllama(); err != nil {
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(heartbeatResponse{Ollama: "error"})
			return
		}
	}

	mu.Lock()
	uptime := time.Since(ollamaStartedAt).Seconds()
	mu.Unlock()

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(heartbeatResponse{Ollama: "running", Uptime: uptime})
}

func main() {
	home, err := os.UserHomeDir()
	if err != nil {
		fatal("cannot find home directory: %v", err)
	}

	certFile := filepath.Join(home, ".ollama", "cert.pem")
	keyFile := filepath.Join(home, ".ollama", "key.pem")

	// Reverse proxy to Ollama
	target, _ := url.Parse(ollamaURL)
	proxy := httputil.NewSingleHostReverseProxy(target)

	// Static file server from embedded dist files
	sub, _ := fs.Sub(staticFiles, "static")
	files := http.FileServer(http.FS(sub))

	// Initialize heartbeat so watchdog doesn't trigger immediately on startup
	lastHeartbeat = time.Now()

	mux := http.NewServeMux()
	mux.HandleFunc("/heartbeat", handleHeartbeat)
	mux.HandleFunc("/v1/", ensureOllama(proxy.ServeHTTP))
	mux.HandleFunc("/api/", ensureOllama(proxy.ServeHTTP))
	mux.Handle("/", files)

	cert, err := tls.LoadX509KeyPair(certFile, keyFile)
	if err != nil {
		fatal("TLS cert error: %v\nRun the install script to generate certificates.", err)
	}

	srv := &http.Server{
		Addr:    listenAddr,
		Handler: mux,
		TLSConfig: &tls.Config{
			Certificates: []tls.Certificate{cert},
		},
	}

	fmt.Printf("Excel AI: https://%s -> Ollama %s (on demand)\n", listenAddr, ollamaURL)
	if err := srv.ListenAndServeTLS("", ""); err != nil {
		fatal("server: %v", err)
	}
}

func fatal(format string, args ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}
