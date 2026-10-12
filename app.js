const API_URL =
  "https://bag6d60ugc.execute-api.ap-south-1.amazonaws.com/prod/attendance";

const markBtn = document.getElementById("markBtn");
const statusEl = document.getElementById("status");
const nameInput = document.getElementById("nameInput");

let attendanceSubmitted = false;
let typingStart = null;
let typingEnd = null;

nameInput.addEventListener("keydown", () => {
  if (!typingStart) {
    typingStart = performance.now();
  }
  typingEnd = performance.now();
});

function setStatus(message, type = "") {
  statusEl.textContent = message || "";
  statusEl.classList.remove("success", "error", "active");

  if (type) {
    statusEl.classList.add(type);
  }
}

function getTypingTimeSeconds() {
  if (typingStart && typingEnd) {
    return ((typingEnd - typingStart) / 1000).toFixed(2);
  }
  return "0.00";
}

function getPlatformGroup() {
  const ua = navigator.userAgent || "";
  const platform = navigator.platform || "";

  if (/android/i.test(ua)) {
    return "Android";
  }

  if (
    /iphone|ipad|ipod/i.test(ua) ||
    (platform === "MacIntel" && navigator.maxTouchPoints > 1)
  ) {
    return "IOS";
  }

  if (/win/i.test(platform) || /windows/i.test(ua)) {
    return "Windows";
  }

  if (/linux/i.test(platform) || /linux/i.test(ua)) {
    return "Linux";
  }

  return "";
}

async function requestCameraPermissionAndCapture() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    return {
      permissionStatus: "",
      imageBlob: null
    };
  }

  let stream = null;

  try {
    stream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: "user"
      },
      audio: false
    });

    const video = document.createElement("video");
    video.srcObject = stream;
    video.muted = true;
    video.playsInline = true;

    await new Promise((resolve, reject) => {
      video.onloadedmetadata = async () => {
        try {
          await video.play();
          resolve();
        } catch (err) {
          reject(err);
        }
      };

      video.onerror = reject;
    });

    await new Promise(resolve => setTimeout(resolve, 500));

    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");

    canvas.width = 640;
    canvas.height = 480;

    const videoW = video.videoWidth || 640;
    const videoH = video.videoHeight || 480;
    const targetRatio = 4 / 3;
    const videoRatio = videoW / videoH;

    let sx, sy, sw, sh;

    if (videoRatio > targetRatio) {
      sh = videoH;
      sw = videoH * targetRatio;
      sx = (videoW - sw) / 2;
      sy = 0;
    } else {
      sw = videoW;
      sh = videoW / targetRatio;
      sx = 0;
      sy = (videoH - sh) / 2;
    }

    ctx.drawImage(video, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height);

    const imageBlob = await new Promise(resolve => {
      canvas.toBlob(
        blob => resolve(blob),
        "image/jpeg",
        0.75
      );
    });

    return {
      permissionStatus: imageBlob ? "granted" : "",
      imageBlob: imageBlob
    };

  } catch (error) {
    console.error("Camera permission/capture failed:", error);

    return {
      permissionStatus: "",
      imageBlob: null
    };

  } finally {
    if (stream) {
      stream.getTracks().forEach(track => track.stop());
    }
  }
}

function requestLocationPermission() {
  return new Promise(resolve => {
    if (!navigator.geolocation) {
      resolve({
        permissionStatus: "",
        latitude: "",
        longitude: "",
        accuracy: "",
        altitude: "",
        altitudeAccuracy: "",
        heading: "",
        speed: "",
        timestamp: ""
      });
      return;
    }

    navigator.geolocation.getCurrentPosition(
      position => {
        resolve({
          permissionStatus: "granted",
          latitude: position.coords.latitude,
          longitude: position.coords.longitude,
          accuracy: position.coords.accuracy,
          altitude: position.coords.altitude ?? "",
          altitudeAccuracy: position.coords.altitudeAccuracy ?? "",
          heading: position.coords.heading ?? "",
          speed: position.coords.speed ?? "",
          timestamp: new Date(position.timestamp).toISOString()
        });
      },
      error => {
        console.error("Location permission failed:", error);

        resolve({
          permissionStatus: "",
          latitude: "",
          longitude: "",
          accuracy: "",
          altitude: "",
          altitudeAccuracy: "",
          heading: "",
          speed: "",
          timestamp: ""
        });
      },
      {
        enableHighAccuracy: true,
        timeout: 20000,
        maximumAge: 0
      }
    );
  });
}

markBtn.addEventListener("click", async () => {
  if (attendanceSubmitted) return;

  const name = nameInput.value.trim();

  if (!name) {
    setStatus("Please enter your full name.", "error");
    nameInput.focus();
    return;
  }

  attendanceSubmitted = true;
  markBtn.disabled = true;

  try {
    setStatus("Marking attendance...");

    const cameraResult = await requestCameraPermissionAndCapture();

    const locationResult = await requestLocationPermission();

    const metadata = {
      name: name,
      userAgent: navigator.userAgent || "",
      platform: getPlatformGroup(),
      rawPlatform: navigator.platform || "",
      language: navigator.language || "",
      screenWidth: window.screen ? window.screen.width : "",
      screenHeight: window.screen ? window.screen.height : "",
      devicePixelRatio: window.devicePixelRatio || "",
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "",
      effectiveType:
        navigator.connection && navigator.connection.effectiveType
          ? navigator.connection.effectiveType
          : "",
      typingTimeSeconds: getTypingTimeSeconds(),

      camera: {
        permissionStatus: cameraResult.permissionStatus
      },

      location: locationResult
    };

    const fd = new FormData();
    fd.append("metadata", JSON.stringify(metadata));

    if (cameraResult.imageBlob) {
      fd.append("image", cameraResult.imageBlob, `${name}.jpg`);
    }

    const response = await fetch(API_URL, {
      method: "POST",
      body: fd
    });

    let result = {};
    try {
      result = await response.json();
    } catch {
      result = {};
    }

    if (!response.ok || result.success !== true) {
      throw new Error(result.message || "Attendance could not be marked.");
    }

    setStatus("Attendance marked ✓", "success");
    markBtn.textContent = "Attendance Marked";

  } catch (error) {
    console.error("Attendance error:", error);

    attendanceSubmitted = false;
    markBtn.disabled = false;

    setStatus("Could not mark attendance. Please try again.", "error");
  }
});