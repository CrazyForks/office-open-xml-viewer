// Interactive 3D hero icon. cube3.glb is a cube with three glowing display
// faces — Front/Right/Top show emissive W / P / E (Word / PowerPoint / Excel).
// Only those three faces are built, so rotation is clamped to a corner view;
// the unfinished back faces are never revealed.
import * as THREE from 'three';
import { GLTFLoader } from 'three/addons/loaders/GLTFLoader.js';

export interface Hero3DOptions {
  rotX?: number;
  rotY?: number;
  dist?: number;
  fov?: number;
}

export function mountHero3D(canvas: HTMLCanvasElement, glbUrl: string, opts: Hero3DOptions = {}): () => void {
  const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const baseRotX = opts.rotX ?? 0.5;
  const baseRotY = opts.rotY ?? -0.66;

  const renderer = new THREE.WebGLRenderer({ canvas, alpha: true, antialias: true });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.outputColorSpace = THREE.SRGBColorSpace;
  renderer.toneMapping = THREE.ACESFilmicToneMapping;
  renderer.toneMappingExposure = 1.05;

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(opts.fov ?? 30, 1, 0.1, 100);
  camera.position.set(0, 0, opts.dist ?? 7.2);

  // Dim environment — the emissive screens carry the image. Soft, broad lights
  // (no tight speculars) so the glass doesn't blow out into hotspots.
  scene.add(new THREE.AmbientLight(0xffffff, 0.34));
  const key = new THREE.DirectionalLight(0xffffff, 0.22);
  key.position.set(-1, 7, 3);
  scene.add(key);
  const rim = new THREE.DirectionalLight(0x9fb4ff, 0.4);
  rim.position.set(-5, -1, 2);
  scene.add(rim);

  const pivot = new THREE.Group();
  scene.add(pivot);

  new GLTFLoader().load(glbUrl, (gltf) => {
    const model = gltf.scene;

    // The glass covers (Glass_NoRefract) have no transmission extension, so
    // three renders them opaque and they hide the glowing letters behind them.
    // Make them a faint glassy sheen, and push the emissive W/P/E letters so
    // the neon reads against the dark frame.
    model.traverse((obj) => {
      const mesh = obj as THREE.Mesh;
      const mat = mesh.material as THREE.MeshPhysicalMaterial | undefined;
      if (!mat || !('name' in mat)) return;
      if (mat.name === 'Glass_NoRefract') {
        mat.transparent = true;
        mat.opacity = 0.08;
        mat.roughness = 0.15;
        mat.metalness = 0;
        mat.clearcoat = 0;          // kill the blown-out specular hotspot
        mat.depthWrite = false;
      } else if (mat.name === 'Silver_Frame') {
        mat.roughness = 0.42;       // diffuse the highlight, no mirror hotspot
      } else if (mat.name.startsWith('Glow')) {
        // Keep the file's emissive hue (W=blue, P=orange, E=green); a modest
        // bump + tone mapping keeps the neon bright without clipping to white.
        mat.emissiveIntensity = 1.9;
        mat.toneMapped = true;
      }
    });

    const box = new THREE.Box3().setFromObject(model);
    const size = box.getSize(new THREE.Vector3());
    const center = box.getCenter(new THREE.Vector3());
    model.position.sub(center);
    const maxDim = Math.max(size.x, size.y, size.z) || 1;
    model.scale.setScalar(2.6 / maxDim);
    pivot.add(model);
    canvas.classList.add('ready');
  });

  let targetX = 0;
  let targetY = 0;
  const onPointer = (e: PointerEvent) => {
    const r = canvas.getBoundingClientRect();
    if (!r.width) return;
    const nx = ((e.clientX - r.left) / r.width) * 2 - 1;
    const ny = ((e.clientY - r.top) / r.height) * 2 - 1;
    targetY = nx * 0.26;
    targetX = ny * 0.18;
  };
  window.addEventListener('pointermove', onPointer);

  function resize(): void {
    const w = canvas.clientWidth || canvas.parentElement?.clientWidth || 480;
    const h = canvas.clientHeight || canvas.parentElement?.clientHeight || 480;
    renderer.setSize(w, h, false);
    camera.aspect = w / h;
    camera.updateProjectionMatrix();
  }
  const ro = new ResizeObserver(resize);
  ro.observe(canvas);
  resize();

  let raf = 0;
  let t = 0;
  let curX = 0;
  let curY = 0;
  function frame(): void {
    raf = requestAnimationFrame(frame);
    t += 0.01;
    const swayY = reduce ? 0 : Math.sin(t) * 0.13;
    const swayX = reduce ? 0 : Math.cos(t * 0.8) * 0.05;
    curX += (targetX + swayX - curX) * 0.06;
    curY += (targetY + swayY - curY) * 0.06;
    pivot.rotation.x = baseRotX + curX;
    pivot.rotation.y = baseRotY + curY;
    renderer.render(scene, camera);
  }
  frame();

  return () => {
    cancelAnimationFrame(raf);
    window.removeEventListener('pointermove', onPointer);
    ro.disconnect();
    renderer.dispose();
  };
}
