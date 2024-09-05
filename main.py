import re

from openai import OpenAI


def sanitize(string):
    return string.replace('"', '')


class Candidate:
    def __init__(self, sams_name, choice_range):
        self.name = sams_name
        self.choice_rank = {}

        for rank in choice_range:
            self.choice_rank[rank] = 0
        self.rank_size = len(self.choice_rank)
        self.borda_count = 0
        self.reasons = list()
        self.examples = list()

    def set_choice(self, choice, reason, examples):
        try:
            count = self.choice_rank[choice]
            self.choice_rank[choice] = count + 1
        except KeyError:
            self.choice_rank[choice] = 1

        self.borda_count = sum((value * (self.rank_size+1 - key)) for key, value in self.choice_rank.items())

        self.reasons.append(sanitize(reason))
        self.examples.append(sanitize(examples))

    def get_borda_count(self):
        return self.borda_count

    def get_reasons(self):
        return '.'.join(self.reasons)

    def get_examples(self):
        return '.'.join(self.examples)

    def __gt__(self, other):
        if self.get_borda_count() < other.get_borda_count():
            return True
        else:
            return False

    def __eq__(self, other):
        return self.get_borda_count() == other.get_borda_count()

    def __repr__(self):
        choice_results = ' '.join([str(f'{k}:{v}') for k, v in self.choice_rank.items()])
        return f'{self.name} {choice_results} Borda Count: {self.borda_count}'

    def get_counts(self):
        choice_results = '\t'.join([str(f'{v}') for k, v in self.choice_rank.items()])
        return f'{choice_results}\t{self.borda_count}'

def parse_spreadsheet(sheet_file):
    rows = list()

    with open(sheet_file, 'r') as f:
        new_row = None
        header_done = False
        for line in f.readlines():
            if not header_done:
                header_done = True
            else:
                # if the row doesn't start with a timestamp (and it's not the header row) it must be a
                # continuation of the previous row
                m = re.search(r"^\d{,2}/\d{,2}/\d{,4} \d{,2}:\d{,2}:\d{,2}\t", line.strip())
                if m is not None:
                    if new_row is not None:
                        rows.append(new_row)
                    new_row = line.strip()
                else:
                    new_row += line.strip()
        rows.append(new_row)
    return rows


def extract_data(sheet_rows):
    the_candidates = {}
    the_unsuitables = {}
    suggestions = []

    for row_num, row in enumerate(sheet_rows):
        # print(f'Processing {row_num}')
        parts = row.split('\t')
        # collect the 4 top choices, reasons and examples
        for idx, col in enumerate([1, 4, 7, 10]):
            name = parts[col].strip()
            if len(name) > 0:
                if name.startswith('5th choice'):
                    fifth = name.split(' ')
                    name = f'{fifth[-2]} {fifth[-1]}'
                    choice_num = 5
                else:
                    choice_num = idx + 1
                if name not in the_candidates:
                    candidate = Candidate(name, range(1, 6))
                else:
                    candidate = the_candidates[name]

                candidate.set_choice(choice_num, parts[col + 1], parts[col + 2])
                the_candidates[name] = candidate

        # collect the unsuitable candidates and reasons
        for idx, no_gos in enumerate([13, 15]):
            no_go = parts[no_gos].strip()
            if no_go != '' and no_go.lower() != 'none':
                if no_go not in the_unsuitables:
                    unsuitable = Candidate(no_go, range(1, 3))
                else:
                    unsuitable = the_unsuitables[no_go]
                unsuitable.set_choice(idx + 1, parts[no_gos + 1], '')
                the_unsuitables[no_go] = unsuitable

        suggestion = parts[17].strip()
        if suggestion != '':
            suggestions.append(suggestion)


    return (dict(sorted(the_candidates.items(), key=lambda item: item[1].get_borda_count(), reverse=True)),
            dict(sorted(the_unsuitables.items(), key=lambda item: item[1].get_borda_count(), reverse=True)),
            suggestions)


if __name__ == '__main__':
    spreadsheet = 'C:/Users/Andrew.Tuley/Downloads/Harrogate Samaritans - New Director Selection - 2024 (Responses) - Form responses 1.tsv'
    rows = parse_spreadsheet(spreadsheet)

    sorted_candidates, unsuitable_candidates, branch_suggestions = extract_data(rows)
    # for c in sorted_candidates:
    #     print(sorted_candidates[c])
    #
    for u in unsuitable_candidates:
        print(unsuitable_candidates[u])

    # print('.'.join(branch_suggestions))

    client = OpenAI()
    top_candidates = [x for x in sorted_candidates.items() if x[1].get_borda_count() > 30]
    content = f'Summarize the feedback about the following candidates for the role of Director. Combine the Qualities and Skills and Examples for each candidate. Produce the output for each candidate in the same format. Use British English spelling.'
    print ('Name\t1st\t2nd\t3rd\t4th\t5th\tBorda Count')
    for top in top_candidates:
        print(f'{top[0]}\t{top[1].get_counts()}')
        content += f'Candidate: {top[0]}: Qualities and Skills:"' + ''.join(
            top[1].get_reasons()) + '" Examples: "' + ''.join(
            top[1].get_examples()) + '"'

    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": f"{content}",
            }
        ]
    )
    print(completion.choices[0].message.content)

    content = f'Summarize the following suggestions for branch improvements. Attempt to find themes where possible and use British English spelling. "{'.'.join(branch_suggestions)}"'
    # print(content)
    completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": f"{content}",
            }
        ]
    )
    print(completion.choices[0].message.content)
