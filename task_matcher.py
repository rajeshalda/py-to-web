import requests
from typing import Dict, List, Optional, Any
from config import AZURE_OPENAI_KEY, AZURE_OPENAI_ENDPOINT
from utils import clean_text

class TaskMatcher:
    def __init__(self):
        self.task_context = {}
    
    def build_task_context(self, tasks: List[Dict[str, Any]], intervals_client) -> Dict[str, Dict]:
        """Build context for task matching."""
        print("Building task context...")
        task_context = {}
        
        for task in tasks:
            print(f"  Processing: {task['title']}")
            task_id = task['id']
            
            task_context[task_id] = {
                'title_words': [],
                'project_name': "",
                'module_name': "",
                'keywords': [],
                'usage_count': 0
            }
            
            # Process title
            clean_title = clean_text(task['title'], remove_emoji=True)
            title_words = [
                word.lower() for word in clean_title.split()
                if len(word) > 2 and not word.isdigit()
            ]
            
            task_context[task_id]['title_words'] = title_words
            task_context[task_id]['keywords'].extend(title_words)
            
            # Add project context
            try:
                if task.get('projectid'):
                    project_info = intervals_client.get_project(task['projectid'])
                    if project_info and project_info.get('name'):
                        task_context[task_id]['project_name'] = project_info['name']
                        
                        project_words = clean_text(project_info['name'], remove_emoji=True)
                        project_words = [
                            word.lower() for word in project_words.split()
                            if len(word) > 2 and not word.isdigit()
                        ]
                        task_context[task_id]['keywords'].extend(project_words)
            except Exception as e:
                print(f"    Unable to fetch project info for task {task_id}: {str(e)}")
            
            # Remove duplicates from keywords
            task_context[task_id]['keywords'] = list(set(task_context[task_id]['keywords']))
        
        self.task_context = task_context
        return task_context
    
    def direct_match(self, meeting_title: str, tasks: List[Dict[str, Any]]) -> Optional[str]:
        """Find direct keyword matches between meeting title and tasks."""
        clean_title = clean_text(meeting_title, remove_emoji=True)
        title_words = [
            word.lower() for word in clean_title.split()
            if len(word) > 2 and not word.isdigit()
        ]
        
        best_match = {
            'task_id': None,
            'score': 0
        }
        
        for task in tasks:
            task_id = task['id']
            task_keywords = self.task_context[task_id]['keywords']
            
            match_count = sum(1 for word in title_words if word in task_keywords)
            
            if title_words:
                score = match_count / len(title_words)
                
                if score > best_match['score']:
                    best_match['task_id'] = task_id
                    best_match['score'] = score
        
        if best_match['score'] > 0.5:
            print(f"Direct match found (Score: {round(best_match['score'], 2)})")
            return best_match['task_id']
        
        return None
    
    def ai_match(self, meeting_subject: str, tasks: List[Dict[str, Any]]) -> Optional[str]:
        """Use Azure OpenAI to match meeting to task."""
        clean_subject = clean_text(meeting_subject, remove_emoji=True)
        
        # Build task analysis string
        task_analysis = "\n".join([
            f"Task ID: {task['id']}\n"
            f"Title: {task['title']}\n"
            f"Project: {self.task_context[task['id']]['project_name']}\n"
            f"Keywords: {', '.join(self.task_context[task['id']]['keywords'])}\n"
            for task in tasks
        ])
        
        prompt = f"""
Match this meeting title with the most relevant task.

Meeting Title: {clean_subject}

Available Tasks:
{task_analysis}

Instructions:
1. Match based on meeting title and task titles/keywords
2. Look for direct keyword matches first
3. Consider project context
4. For infrastructure/network meetings, prefer infrastructure tasks
5. For client-specific meetings, match to respective tasks

Response format:
{{
  "taskId": "numeric_id_or_NO_MATCH",
  "confidence": "high|medium|low"
}}
"""
        
        try:
            headers = {
                "api-key": AZURE_OPENAI_KEY,
                "Content-Type": "application/json"
            }
            
            body = {
                "messages": [
                    {
                        "role": "system",
                        "content": "You are a task matcher focusing on title keywords and project context."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "temperature": 0.3,
                "max_tokens": 100
            }
            
            response = requests.post(
                AZURE_OPENAI_ENDPOINT,
                headers=headers,
                json=body
            )
            response.raise_for_status()
            
            result = response.json()
            match_result = result["choices"][0]["message"]["content"]
            match_data = eval(match_result)  # Safe in this context as we control the input
            
            print(f"AI Match Result:")
            print(f"  Task ID: {match_data['taskId']}")
            print(f"  Confidence: {match_data['confidence']}")
            
            if match_data['confidence'] == "low":
                return None
                
            return match_data['taskId']
            
        except Exception as e:
            print(f"Error in AI matching: {str(e)}")
            return None 