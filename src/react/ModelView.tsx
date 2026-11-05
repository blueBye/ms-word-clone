"use client";
import { useEffect, useRef, useState } from "react";
import type { DrawingInline } from "../model/document.js";
import { modelData } from "../core/assets.js";
export function ModelView({
  model,
  onReady,
}: {
  model: NonNullable<DrawingInline["model3d"]>;
  onReady?: (
    poster: string,
    size: number,
    center: [number, number, number],
  ) => void;
}) {
  const root = useRef<HTMLDivElement>(null),
    ready = useRef(onReady);
  ready.current = onReady;
  const [error, setError] = useState("");
  useEffect(() => {
    if (root.current?.closest('[aria-hidden="true"]')) return;
    let disposed = false,
      cleanup = () => {};
    void (async () => {
      try {
        const bytes = modelData(model.dataUrl),
          [T, { GLTFLoader }, { OrbitControls }] = await Promise.all([
            import("three"),
            import("three/addons/loaders/GLTFLoader.js"),
            import("three/addons/controls/OrbitControls.js"),
          ]);
        const manager = new T.LoadingManager();
        manager.setURLModifier((url) => {
          if (!/^(data:|blob:)/.test(url))
            throw new Error(
              "External model assets are not supported. Embed them in the GLB file.",
            );
          return url;
        });
        const gltf = await new GLTFLoader(manager).parseAsync(
          bytes.buffer as ArrayBuffer,
          "",
        );
        if (disposed) return;
        const host = root.current!,
          renderer = new T.WebGLRenderer({
            alpha: true,
            antialias: true,
            preserveDrawingBuffer: true,
          });
        renderer.setPixelRatio(Math.min(devicePixelRatio, 2));
        renderer.setSize(host.clientWidth || 400, host.clientHeight || 250);
        host.appendChild(renderer.domElement);
        renderer.domElement.style.width = "100%";
        renderer.domElement.style.height = "100%";
        const scene = new T.Scene(),
          box = new T.Box3().setFromObject(gltf.scene),
          center = box.getCenter(new T.Vector3()),
          size = Math.max(0.001, ...box.getSize(new T.Vector3()).toArray());
        const camera = new T.PerspectiveCamera(
          40,
          (host.clientWidth || 400) / (host.clientHeight || 250),
          size / 1000,
          size * 100,
        );
        camera.position.set(
          center.x + size * 0.2,
          center.y + size * 0.1,
          center.z + size * 2,
        );
        camera.lookAt(center);
        gltf.scene.rotation.set(
          ((model.rotation?.x ?? 0) * Math.PI) / 180,
          ((model.rotation?.y ?? 0) * Math.PI) / 180,
          ((model.rotation?.z ?? 0) * Math.PI) / 180,
        );
        scene.add(gltf.scene);
        scene.add(new T.HemisphereLight(0xffffff, 0x555566, 3));
        const light = new T.DirectionalLight(0xffffff, 3);
        light.position.set(size, size, size * 2);
        scene.add(light);
        const controls = new OrbitControls(camera, renderer.domElement);
        controls.target.copy(center);
        controls.enablePan = false;
        controls.update();
        const render = () => renderer.render(scene, camera);
        controls.addEventListener("change", render);
        render();
        ready.current?.(
          renderer.domElement.toDataURL("image/png"),
          size,
          center.toArray() as [number, number, number],
        );
        const observer = new ResizeObserver(() => {
          const w = host.clientWidth,
            h = host.clientHeight;
          if (w && h) {
            renderer.setSize(w, h);
            camera.aspect = w / h;
            camera.updateProjectionMatrix();
            render();
          }
        });
        observer.observe(host);
        cleanup = () => {
          observer.disconnect();
          controls.dispose();
          gltf.scene.traverse((o) => {
            if (o instanceof T.Mesh) {
              o.geometry.dispose();
              const materials = Array.isArray(o.material)
                ? o.material
                : [o.material];
              materials.forEach((m) => {
                for (const value of Object.values(m))
                  if (value instanceof T.Texture) value.dispose();
                m.dispose();
              });
            }
          });
          renderer.dispose();
          renderer.domElement.remove();
        };
      } catch (error) {
        if (!disposed)
          setError(
            error instanceof Error
              ? error.message
              : "Could not display this model.",
          );
      }
    })();
    return () => {
      disposed = true;
      cleanup();
    };
  }, [model.dataUrl, model.rotation?.x, model.rotation?.y, model.rotation?.z]);
  return (
    <div
      ref={root}
      style={{ width: "100%", height: "100%", position: "relative" }}
    >
      {error && <span role="status">{error}</span>}
    </div>
  );
}
