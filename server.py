import re
import requests
from openai import OpenAI
from flask import Flask, request, jsonify
from PdfToPpt.index import generate_presentation
# Initialize Flask app
app = Flask(__name__)
# generate_ppt_endpoint(app)
# --- Helper Function to Parse Bitbucket URL ---
def parse_pr_url(url):
    """Extracts workspace, repo_slug, and pr_id from a Bitbucket URL."""
    pattern = r"bitbucket.org/([^/]+)/([^/]+)/pull-requests/(\d+)"
    match = re.search(pattern, url)
    if match:
        workspace, repo_slug, pr_id = match.groups()
        return workspace, repo_slug, pr_id
    return None, None, None

# --- Core Logic Functions ---
def get_bitbucket_diff(workspace, repo_slug, pr_id, token):
    diff_url = f"https://api.bitbucket.org/2.0/repositories/{workspace}/{repo_slug}/pullrequests/{pr_id}/diff"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(diff_url, headers=headers)
        response.raise_for_status()
        return response.text, None
    except requests.exceptions.RequestException as e:
        return None, str(e)

def get_openai_analysis(diff_text, system_prompt, api_key):
    try:
        client = OpenAI(api_key=api_key)
        completion = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Here is the code diff:\n\n```diff\n{diff_text}\n```"}
            ]
        )
        return completion.choices[0].message.content, None
    except Exception as e:
        return None, str(e)

def post_bitbucket_comment(workspace, repo_slug, pr_id, comment_text, token):
    comment_url = f"https://api.bitbucket.org/2.0/repositories/{workspace}/{repo_slug}/pullrequests/{pr_id}/comments"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"content": {"raw": comment_text}}

    try:
        response = requests.post(comment_url, headers=headers, json=payload)
        response.raise_for_status()
        return True, None
    except requests.exceptions.RequestException as e:
        return False, str(e)

# --- Shared Review Logic ---
def handle_review(pr_url, token, api_key, action):
    prompts = {
        "review": "Act as a senior software engineer and perform a code review. Look for bugs, style issues, and potential improvements.",
        "analyse": "Perform a deep analysis of these code changes, considering performance, security, and maintainability.",
        "describe": "Provide a high-level summary of the following code changes."
    }

    # Parse PR URL
    workspace, repo_slug, pr_id = parse_pr_url(pr_url)
    if not workspace:
        return jsonify({"error": "Invalid Bitbucket PR URL format."}), 400

    # Get diff
    diff, error = get_bitbucket_diff(workspace, repo_slug, pr_id, token)
    if error:
        return jsonify({"error": f"Failed to fetch diff: {error}"}), 502

    # Get analysis
    analysis, error = get_openai_analysis(diff, prompts[action], api_key)
    if error:
        return jsonify({"error": f"Failed to get analysis from OpenAI: {error}"}), 502

    # Post back to PR
    header = f"### 🤖 AI {action.capitalize()} Result\n\n---\n\n"
    success, error = post_bitbucket_comment(workspace, repo_slug, pr_id, header + analysis, token)
    if not success:
        return jsonify({"error": f"Failed to post comment to Bitbucket: {error}"}), 502

    return jsonify({"status": "success", "message": f"{action.capitalize()} posted to PR #{pr_id}."}), 200

# --- Endpoints ---
@app.route("/review", methods=["POST"])
def review_endpoint():
    bb_token = request.headers.get("bb-repo-access-token")
    openai_key = request.headers.get("openai-api-key")
    data = request.get_json()
    pr_url = data.get("pr_url")
    return handle_review(pr_url, bb_token, openai_key, "review")

@app.route("/analyse", methods=["POST"])
def analyse_endpoint():
    bb_token = request.headers.get("bb-repo-access-token")
    openai_key = request.headers.get("openai-api-key")
    data = request.get_json()
    pr_url = data.get("pr_url")
    return handle_review(pr_url, bb_token, openai_key, "analyse")

@app.route("/describe", methods=["POST"])
def describe_endpoint():
    bb_token = request.headers.get("bb-repo-access-token")
    openai_key = request.headers.get("openai-api-key")
    data = request.get_json()
    pr_url = data.get("pr_url")
    return handle_review(pr_url, bb_token, openai_key, "describe")

@app.route("/generate-ppt", methods=["POST"])
def generate_ppt_endpoint():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "JSON body required"}), 400

        slide_count = data.get("slide_count")
        summary = data.get("summary")

        if not slide_count or not isinstance(slide_count, int) or slide_count < 1:
            return jsonify({"error": "Valid slide_count (int > 0) required"}), 400

        if not summary or not isinstance(summary, str) or summary.strip() == "":
            return jsonify({"error": "Non-empty summary string required"}), 400

        ppt_url = generate_presentation(slide_count, summary)

        return jsonify({
            "message": "Presentation generated successfully",
            "presentation_url": ppt_url
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Run app
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
