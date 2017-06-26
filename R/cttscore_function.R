#' A function to score the responses to get dichotomous score, in order to do further analysis. 
#' 
#' @param response The response matrix, each row is an examinee, each column is one item.
#' @param key The answer key, either one value for all the items, or a vector in which each value for one item
#' @return A sum score vector and a scored matrix
#' @export 
#' @examples
#' ctt_Score <- function(response=response_matrix, key=1, na=NA)

ctt_Score <- function(response, key)
{
    if (length(key)==ncol(response))
    { 
      scored <- t(apply(response, 1, function(x){ifelse(x==(key),1,0)}))
    } else {
      if (length(key)==1) 
        {
        scored <-1*(response==key)
        } else stop("Number of items is not equal to the length of key.")
    }
  
  score<-rowSums(scored)
  list(score=score,scored=scored)
}
