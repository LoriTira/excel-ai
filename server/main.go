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
	"strings"
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
	mu              sync.Mutex
	ollamaCmd       *exec.Cmd
	ollamaRunning   bool
	ollamaStarting  bool
	startDone       chan struct{} // closed when startup attempt completes
	lastHeartbeat   time.Time
	ollamaStartedAt time.Time
	ollamaPath      string // resolved path to ollama binary
)

// findOllama resolves the full path to the ollama binary.
// The install script writes the path to ollama-path next to the server binary.
// Falls back to common locations if that file is missing.
func findOllama() string {
	// 1. Read path saved by install script (most reliable)
	if exePath, err := os.Executable(); err == nil {
		configPath := filepath.Join(filepath.Dir(exePath), "ollama-path")
		if data, err := os.ReadFile(configPath); err == nil {
			p := strings.TrimSpace(string(data))
			if p != "" {
				if _, err := os.Stat(p); err == nil {
					return p
				}
			}
		}
	}

	// 2. Check PATH (works in dev, rarely in production under launchd/Scheduled Task)
	if p, err := exec.LookPath("ollama"); err == nil {
		return p
	}

	// 3. Common install locations as last resort
	candidates := []string{"/usr/local/bin/ollama", "/opt/homebrew/bin/ollama"}
	if runtime.GOOS == "windows" {
		candidates = nil
		if localAppData := os.Getenv("LOCALAPPDATA"); localAppData != "" {
			candidates = append(candidates, filepath.Join(localAppData, "Programs", "Ollama", "ollama.exe"))
		}
	} else {
		if home, err := os.UserHomeDir(); err == nil {
			candidates = append(candidates, filepath.Join(home, ".local", "bin", "ollama"))
		}
	}
	for _, p := range candidates {
		if _, err := os.Stat(p); err == nil {
			return p
		}
	}
	return "ollama"
}

func startOllama() error {
	mu.Lock()

	if ollamaRunning {
		mu.Unlock()
		return nil
	}

	// Another goroutine is already starting Ollama — wait for it
	if ollamaStarting {
		ch := startDone
		mu.Unlock()
		<-ch
		mu.Lock()
		ok := ollamaRunning
		mu.Unlock()
		if ok {
			return nil
		}
		return fmt.Errorf("ollama startup failed (concurrent)")
	}

	// Check if Ollama is already running externally
	if isOllamaReachable() {
		ollamaRunning = true
		ollamaStartedAt = time.Now()
		go watchdog()
		mu.Unlock()
		fmt.Println("Ollama already running.")
		return nil
	}

	// Mark as starting so concurrent callers wait
	ollamaStarting = true
	startDone = make(chan struct{})

	// Spawn Ollama
	fmt.Printf("Starting Ollama (%s)...\n", ollamaPath)
	cmd := exec.Command(ollamaPath, "serve")
	cmd.Stdout = nil
	cmd.Stderr = nil
	if err := cmd.Start(); err != nil {
		ollamaStarting = false
		close(startDone)
		mu.Unlock()
		return fmt.Errorf("failed to start ollama (%s): %w", ollamaPath, err)
	}

	ollamaCmd = cmd
	mu.Unlock()

	// Poll until Ollama is reachable (no lock held — allows heartbeats/requests to queue)
	err := waitForOllama()

	mu.Lock()
	ollamaStarting = false
	if err != nil {
		cmd.Process.Kill()
		cmd.Wait()
		ollamaCmd = nil
		close(startDone)
		mu.Unlock()
		return err
	}

	ollamaRunning = true
	ollamaStartedAt = time.Now()
	close(startDone)
	go watchdog()
	mu.Unlock()
	fmt.Println("Ollama started.")
	return nil
}

func stopOllama() {
	mu.Lock()
	defer mu.Unlock()

	if !ollamaRunning {
		return
	}

	// Kill the process we spawned, if any
	if ollamaCmd != nil && ollamaCmd.Process != nil {
		if runtime.GOOS == "windows" {
			ollamaCmd.Process.Kill()
		} else {
			ollamaCmd.Process.Signal(os.Interrupt)
		}
		ollamaCmd.Wait()
	}

	// Also kill any Ollama process we didn't spawn (e.g. started by Ollama's own auto-start)
	killAllOllama()

	ollamaCmd = nil
	ollamaRunning = false
	fmt.Println("Ollama stopped.")
}

// killAllOllama terminates all Ollama processes on the system.
func killAllOllama() {
	if runtime.GOOS == "windows" {
		exec.Command("taskkill", "/IM", "ollama.exe", "/F").Run()
	} else {
		exec.Command("pkill", "-x", "ollama").Run()
	}
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
		if err := startOllama(); err != nil {
			http.Error(w, fmt.Sprintf("Failed to start Ollama: %v", err), http.StatusBadGateway)
			return
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
	mu.Unlock()

	if err := startOllama(); err != nil {
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(heartbeatResponse{Ollama: "error"})
		return
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

	// Resolve ollama binary path at startup
	ollamaPath = findOllama()
	fmt.Printf("Ollama binary: %s\n", ollamaPath)

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
