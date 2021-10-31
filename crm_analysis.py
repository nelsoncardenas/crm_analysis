import json

from modules.team_evaluator import TeamEvaluator

if __name__ == "__main__":
    with open("conf.json", "r") as file:
        conf = json.load(file)

    evaluator = TeamEvaluator(conf)
    evaluator.evaluate()
