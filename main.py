import os
import time

import word
from loguru import logger

path = './Articles/'  # Word files folder

if __name__ == '__main__':
    try:
        docs = os.listdir(path)
        if not docs:
            raise Exception('No files in the folder')
    except Exception as e:
        logger.error(e)
        time.sleep(2)
        input('Press Any Key To Exit')
        exit()

    instruction = []
    total = 0
    for file_name in docs:
        logger.debug(f'Processing {file_name}')

        file_name = os.path.join(path, file_name)
        filePath = os.path.abspath(file_name)

        amount = word.check(filePath)  # Get words number

        logger.success(f'{file_name} done, {amount} words')

        instruction.append(f'{file_name}')
        instruction.append(f'-words: {amount}')

        total += amount  # words number for all files

    instruction.append(f'Words of all files: {total}')

    # write the instruction to a file
    with open('./result.txt', 'w', encoding='utf-8') as f:
        for line in instruction:
            f.write(line + '\n')

    logger.success('Result all saved to result.txt')

    time.sleep(2)

    # Press Any Key To Exit
    input('Press Any Key To Exit')
    exit()
