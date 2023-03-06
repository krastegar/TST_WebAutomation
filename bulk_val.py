#-------------------------------------------------------------------------------------------
# 
#   Purpose:
#   Creating a program that will be able to take in RangeExports and IncomingMessageMonitor 
#   Exports and automate making summary reports of the export files. 
#    
#   Author: Kiarash Rastegar
#   Date: 2/27/23
#-------------------------------------------------------------------------------------------

import os
import sys 
import pandas as pd
import numpy as np
import gensim
from sklearn.metrics.pairwise import cosine_similarity
from util import (
    tstRangeQuery_demographic,
    tstRangeQuery_lab, 
    completeness_report, 
    database_connection,
    combined_MM
)
def main():
    # connecting python to microsoft access file
    # Need to input name of .accdb file to run pipe line
    folder_path = r'C:\Users\krastega\Desktop\MicrosoftAccessDB' # will be used as an input 
    
    # only need 1 connection for TST .accdb file 
    conn, cursor = database_connection(folder_path=folder_path,
                                       file_name='TST_DIE_02132023_02212023_exp022323')

    # looking at Disease Incidents table specifically and reading it into a dataframe
    # Also making completeness report for Disease Incident Export table 
    query = tstRangeQuery_demographic(test_center_1='Palomar', test_center_2='Pomerado')
    df_demo = completeness_report(query=query, conn=conn)

    # Looking at Laborartory Systems table and making completeness report
    query = tstRangeQuery_lab(test_center_1='Palomar', test_center_2='Pomerado')
    df_lab = completeness_report(query=query, conn=conn)

    # take incoming message monitors and produce a list of entries 
    # that were found in Production and not in TST
    prod_df = combined_MM(folder_path=folder_path,
                          startswith='PROD')
    tst_df = combined_MM(folder_path=folder_path,
                          startswith='TST')
    # matching message monitor
    # I want to match on these columns(prod_df[['DILR_ResultTest', 'DILR Accession']])
    # need to mess with TST env accession numbers to get rid of leading 0000

    # inner join on Accession number to get matching from both datasets 
    merged_df = pd.merge(
         pd.DataFrame(tst_df[['DILR_AccessionNumber', 'DILR_ResultedTest']]), 
         pd.DataFrame(prod_df[['DILR_AccessionNumber', 'DILR_ResultedTest']]), 
         on=['DILR_AccessionNumber'], 
         how='inner'
         )
    similarity_df = word2vec(
         merged_df, 
         'DILR_ResultedTest_x', 
         'DILR_ResultedTest_y'
         )
    similarity_df.to_excel('matchv1.xlsx')
    return 
    
def word2vec(df, column1, column2):
    df[column1] = df[column1].astype(str)
    df[column2] = df[column2].astype(str)
    # Preprocess the strings in the X and Y columns
    df['X_processed'] = df[column1].apply(lambda x: [word.lower() for word in x.split()]) #if word.isalpha()])
    df['y_processed'] = df[column2].apply(lambda x: [word.lower() for word in x.split()]) #if word.isalpha()])

    # Concatenate the processed strings into a list of sentences
    sentences = list(df['X_processed']) + list(df['y_processed'])

    # Train a word2vec model on the sentences
    model = gensim.models.Word2Vec(sentences, vector_size=100, window=5, min_count=5, workers=4)

    # Calculating word vectors for each word in the list 
    df['X_vector'] = df['X_processed'].apply(lambda x: np.mean([model.wv[word] for word in x if word in model.wv.key_to_index], axis=0))
    df['Y_vector'] = df['y_processed'].apply(lambda x: np.mean([model.wv[word] for word in x if word in model.wv.key_to_index], axis=0))

    # Compute the cosine similarity between the vectors for each pair of strings
    df['cosine_sim'] = df.apply(lambda row: max(
        cosine_similarity([row['X_vector']], [row['Y_vector']])[0][0],
        cosine_similarity([row['Y_vector']], [row['X_vector']])[0][0]
    ) if not np.isnan(row['X_vector']).any() and not np.isnan(row['Y_vector']).any() else np.nan, axis=1)
    
    df.drop_duplicates(['DILR_ResultedTest_x', 'DILR_ResultedTest_y'], inplace=True)
    new_matches = []
    for (tst, prod) in zip(
        df[column1],
        df[column2]
    ):
        test_df = df.loc[df[column1]==tst] # filtering  based on specific test result
        max_sim = test_df.loc[test_df['cosine_sim'].idxmax()] # finding the row with maximum cosine similarity
        max_sim_score = max_sim['cosine_sim'] # max sim score for specific tst test
        pair_tests = df.loc[(df[column1]==tst) & (df[column2]==prod)] # filter pased on pair of test
        pair_score = pair_tests['cosine_sim'].values[0] # get the score on those pair of tests
        # print(f'Max Score: {max_sim_score}\n Pair Score: {pair_score}')
        if max_sim_score == pair_score:
            print(f"MATCH: {tst} ----- {prod}")
            new_matches.append(df.loc[(df[column1]==tst) & 
                                       (df[column2]==prod) & 
                                       (df['cosine_sim']==pair_score)]
                            )
        else:
            
            prod_test_df = df.loc[df[column2]==prod] # filtering df based on production resultedTest
            prod_max_sim = prod_test_df.loc[prod_test_df['cosine_sim'].idxmax()] # finding row in filtered data that has the highest similarity score
            prod_test_max = prod_max_sim['cosine_sim']# finding the max sim score in second prod df that and seeing if it is larger than original pair sim score. If it is we skip and then continue. If its not we append original pair
            # prod_test_score = None

            if prod_test_max >= pair_score:
                print(
                    f"""
                    Checking to see which test pairs are seen in my condition
                    TST: {tst}
                    PROD: {prod}
                    OG_Pair: {pair_score}
                    ProdTest Check: {prod_test_max}
                    """
                )
                if not all(
                    df.loc[
                    (df[column1]==tst) & 
                    (df[column2]==prod) & 
                    (df['cosine_sim']==pair_score)
                ].isin(new_matches)):
                    new_matches.append(df.loc[(df[column1]==tst) & 
                            (df[column2]==prod) & 
                            (df['cosine_sim']==pair_score)])
                else:
                    continue
            else: 
                raise AssertionError(
                    f"""
                    There has been an issue with the following pairs of tests 
                    in my condition of inequality of similarity scores
                    TST: {tst}
                    PROD: {prod}
                    OG_Pair: {pair_score}
                    ProdTest Check: {prod_test_max}
                    """
                    )
            
    new_df = pd.concat(new_matches)
    return new_df

if __name__=="__main__":
    sys.exit(main())