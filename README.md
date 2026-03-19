<h1>A.S.O.R.A. (Automated Storage & Optimization Re Assignement)</h1>
<p><b>A High-Performance RPA Engine for ERP Integration & Data Transformation</b></p>

<h3>🚀 Overview</h3>
<p>A.S.O.R.A. is a custom-built Robotic Process Automation (RPA) solution designed to bridge the gap between Microsoft Excel and legacy ERP systems (such as Reflex and SAP). By leveraging the Windows API, this tool transforms manual, high-volume data entry tasks into a fully automated, background-monitored process.</p>

<h3>Key Business Impact</h3>
<dl>
  <dt><b>Efficiency</b></dt>
  <dd>Achieved a 70%+ increase in operational processing speed.</dd>
  
  <dt><b>Time Savings</b></dt>
  <dd>Automated 20+ hours of manual data handling per week.</dd>
  
  <dt><b>Accuracy</b></dt>
  <dd>Eliminated human error in SKU de-assignment and bin-to-bin stock movements.</dd>
</dl>

<h3>🛠 Technical Stack</h3>
<dl>
  <dt><b>Language</b></dt>
  <dd>VBA (Visual Basic for Applications)</dd>
  
  <dt><b>Architecture</b></dt>
  <dd>Modular Design with 64-bit Windows API Integration.</dd>
  
  <dt><b>APIs Utilized</b></dt>
  <dd><b>User32.dll:</b> Low-level mouse event control and cursor positioning.</dd>
  <dd><b>Kernel32.dll:</b> High-precision timing and process "Sleep" management.</dd>
  <dd><b>Winmm.dll:</b> Multimedia integration for real-time audio status updates.</dd>
  <dd><b>SAPI (Speech API):</b> Text-to-Speech (TTS) engine for asynchronous monitoring.</dd>
</dl>

<h3>📂 Module Breakdown</h3>

<p><b>1. Automated Location Management (The "Location Bot")</b></p>
<dl>
  <dd><b>File:</b> Module1.bas</dd>
  <dd><b>Function:</b> Uses <code>SetCursorPos</code> and <code>mouse_event</code> to interact with non-native UI elements within the ERP.</dd>
  <dd><b>Features:</b> Includes a <code>WhereIsMyMouse</code> utility to capture screen coordinates, allowing for rapid reconfiguration across different resolutions.</dd>
</dl>

<p><b>2. SKU-Level Batch Processing</b></p>
<dl>
  <dd><b>File:</b> Module2.bas</dd>
  <dd><b>Function:</b> Synchronizes complex Excel datasets with the "Project Pick Location Manager."</dd>
  <dd><b>Features:</b> Implemented an <code>GetAsyncKeyState</code> fail-safe, allowing users to immediately kill the process with the <b>ESC</b> key during emergency interruptions.</dd>
</dl>

<p><b>3. LARA (Logistics Automated Relocation Assistant)</b></p>
<dl>
  <dd><b>File:</b> Module3.bas</dd>
  <dd><b>Function:</b> Specialized engine for "Bin-to-Bin" transfers.</dd>
  <dd><b>Features:</b> Features a real-time <b>Transaction Audit Logger</b> that timestamps every move, ensuring 100% compliance with inventory audit standards.</dd>
</dl>

<p><b>4. Audio & Notification Engine</b></p>
<dl>
  <dd><b>File:</b> Module5.bas</dd>
  <dd><b>Function:</b> Provides a professional UX for background automation.</dd>
  <dd><b>Features:</b> Uses <code>mciSendString</code> to play status alerts and a TTS engine to verbally update the operator on process completion.</dd>
</dl>

<h3>⚙️ Implementation & Usage</h3>
<dl>
  <dt><b>Configuration</b></dt>
  <dd>Target coordinates for ERP buttons are mapped directly within an Excel configuration sheet.</dd>
  
  <dt><b>Execution</b></dt>
  <dd>The engine validates the data integrity of the source sheet before initiating UI interaction.</dd>
  
  <dt><b>Monitoring</b></dt>
  <dd>The bot provides visual feedback (cell coloring) and audio cues throughout the lifecycle of the task.</dd>
</dl>

<h3>👤 Author</h3>
<p><b>Pavel Iliev</b><br>
Data Transformation Analyst | BI & Automation Specialist | Stock Data Analyst<br>
<a href="https://www.linkedin.com/in/pavel-iliev-610640155">LinkedIn Profile</a></p>

<p><i><b>Note:</b> This repository contains the source code modules (.bas) for review. For security reasons, sensitive corporate data and specific ERP connection strings have been sanitized.</i></p>
