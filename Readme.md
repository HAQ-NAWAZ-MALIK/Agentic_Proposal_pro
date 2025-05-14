# Federal Proposal Agent System

![Proposal Agent](https://img.shields.io/badge/AI-Proposal%20Generation-blue)
![Python](https://img.shields.io/badge/Python-3.8%20%7C%203.9%20%7C%203.10%20%7C%203.11-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

A sophisticated multi-agent system for analyzing federal solicitations and generating professional proposal outlines. This system streamlines the federal proposal response process by automatically processing complex solicitation documents, extracting requirements, and creating structured proposal documents.
      ![image](https://github.com/user-attachments/assets/ae627b48-0c9c-4cdf-909f-2f2856415acc)

<p align="center">
  <img src="https://via.placeholder.com/800x400?text=Federal+Proposal+Agent" alt="Federal Proposal Agent illustration" width="600"/>
</p>

## 🌟 Features

- **PDF Solicitation Analysis** - Extracts and processes text and tables from federal solicitation PDFs
- **Requirements Identification** - Identifies "shall/must" statements and required capabilities
- **Evaluation Criteria Analysis** - Extracts and prioritizes scoring factors and evaluation criteria
- **Professional DOCX Generation** - Creates properly formatted proposal outlines with advanced styling
- **Compliance Matrices** - Generates traceability matrices linking requirements to response sections
- **Multi-Agent Architecture** - Uses specialized agents for document processing, requirement analysis, and proposal development

## 📋 Table of Contents

- [Installation](#-installation)
- [Usage](#-usage)
- [Architecture](#-architecture)
- [Document Features](#-document-features)
- [Configuration](#-configuration)
- [Troubleshooting](#-troubleshooting)
- [Contributing](#-contributing)
- [License](#-license)

## 🔧 Installation

### Prerequisites

- Python 3.8 or higher
- OpenAI API key

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/HAQ-NAWAZ-MALIK/federal-proposal-agent.git
   cd federal-proposal-agent
   ```

2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Create a `.env` file in the project root with your OpenAI API key:
   ```
   OPENAI_API_KEY=sk-your-actual-openai-key-here
   DEFAULT_MODEL=gpt-4.1-mini
   ```

4. Create directories for your solicitations:
   ```bash
   mkdir solicitations proposal_outputs
   ```

## 🚀 Usage

### Basic Usage

Place your federal solicitation PDF files in the `solicitations` directory, then run:

```bash
python run_proposal_agent.py
```

This will:
1. Scan for solicitation PDF files
2. Prompt you to select which file to process (or process all)
3. Generate a proposal outline document for each solicitation
4. Save the resulting documents in the `proposal_outputs` directory

### Programmatic Usage

```python
from proposal_agent_system import run_proposal_agent

# Process a single solicitation
result = run_proposal_agent(
    solicitation_path="path/to/solicitation.pdf",
    output_path="path/to/proposal_outline.docx"
)

print(result)
```

## 🏗️ Architecture

The Federal Proposal Agent uses a multi-agent architecture:

```
┌─────────────────────────────────┐
│     Proposal Manager Agent      │
│  (Coordinates Development)      │
└─────────────┬───────────────────┘
              │
              ▼
┌─────────────┴───────────────────┐
│                                 │
▼                                 ▼
┌────────────────────┐  ┌────────────────────┐  ┌────────────────────┐
│Document Processing │  │ Requirements       │  │Proposal Development│
│Agent               │  │ Analysis Agent     │  │Agent               │
└────────────────────┘  └────────────────────┘  └────────────────────┘
```

1. **Proposal Manager Agent**: Coordinates the overall workflow and integrates output
2. **Document Processing Agent**: Extracts and structures content from solicitation PDFs
3. **Requirements Analysis Agent**: Identifies requirements and evaluation criteria
4. **Proposal Development Agent**: Creates structured proposal documents with proper formatting

## 📄 Document Features

The generated proposal outlines include:

- **Professional Formatting**: Consistent styles, fonts, and spacing
- **Structured Hierarchy**: Properly numbered sections and subsections
- **Compliance Matrix**: Table mapping solicitation requirements to proposal sections
- **Response Templates**: Placeholder text indicating what content to include
- **Standard Federal Sections**: Common federal proposal sections pre-formatted

## ⚙️ Configuration

You can customize the agent's behavior through environment variables in your `.env` file:

| Variable | Description | Default |
|----------|-------------|---------|
| `DEFAULT_MODEL` | OpenAI model to use | gpt-4.1-mini |
| `MAX_STEPS` | Maximum agent steps | 30 |
| `VERBOSITY_LEVEL` | Log detail level (0-3) | 2 |
| `LOG_LEVEL` | Python logging level | INFO |

## ❓ Troubleshooting

### Common Issues

1. **PDF Extraction Failures**:
   ```
   Error extracting text from PDF: ...
   ```
   
   **Solution**: Ensure your PDF is not scanned/image-based. The system works best with searchable PDFs.

2. **Authentication Error**:
   ```
   AuthenticationError: Error code: 401 - {'error': {'message': 'Incorrect API key provided...
   ```
   
   **Solution**: Verify your OpenAI API key in the `.env` file.

3. **Document Generation Issues**:
   ```
   Error generating DOCX outline: ...
   ```
   
   **Solution**: Check that python-docx is properly installed and the output directory is writable.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📜 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgements

- [SmolaGents](https://github.com/smol-ai/smolagents) for the agent framework
- [python-docx](https://python-docx.readthedocs.io/) for DOCX generation
- [PyPDF](https://pypdf.readthedocs.io/) for PDF processing
- [OpenAI](https://openai.com/) for language model APIs

---

<p align="center">
  Made with ❤️ for streamlining federal proposal development
</p>
