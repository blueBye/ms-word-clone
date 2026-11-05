"use client";
import { useState } from "react";
import { videoUrl } from "../core/assets.js";
export function VideoView({ url, poster }: { url: string; poster?: string }) {
  const [play, setPlay] = useState(false);
  let source: string;
  try {
    source = videoUrl(url);
  } catch {
    return <span>Unsupported video URL</span>;
  }
  return play ? (
    <iframe
      src={source}
      title="Online video"
      allow="autoplay; encrypted-media; picture-in-picture"
      allowFullScreen
      referrerPolicy="no-referrer"
      style={{ width: "100%", height: "100%", border: 0 }}
    />
  ) : (
    <button
      type="button"
      aria-label="Play video"
      onClick={() => setPlay(true)}
      style={{
        width: "100%",
        height: "100%",
        padding: 0,
        position: "relative",
        background: "#29384a",
      }}
    >
      {poster && (
        <img
          src={poster}
          alt="Video thumbnail"
          style={{ width: "100%", height: "100%", objectFit: "contain" }}
        />
      )}
      <span
        style={{
          position: "absolute",
          left: "50%",
          top: "50%",
          transform: "translate(-50%,-50%)",
          fontSize: "36pt",
          color: "white",
        }}
      >
        ▶
      </span>
    </button>
  );
}
