import { mount } from 'svelte';
import App from './App.svelte';
import '@ooxml-framework-examples/shared/example.css';

mount(App, { target: document.getElementById('app') as HTMLElement });
